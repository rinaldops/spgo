#!/usr/bin/env node
'use strict';

/**
 * Standalone probe: does Entra ID currently allow a modern, MFA-capable login
 * to this SharePoint Online site? Speaks the OAuth device-code protocol
 * directly (no msal-node) so any AADSTS error comes through unmodified,
 * instead of msal-node's generic "network" wrapper hiding the real reason.
 *
 * Usage:
 *   node entra-access-probe.js <clientId> [siteUrl] [tenantId]
 *
 * clientId is required - there is no shared app to fall back to. The community
 * "PnP Management Shell" multi-tenant app was deleted by the PnP team on
 * 2024-09-09 for security reasons, so every tenant needs its own App Registration.
 * Fastest way to create one:
 *   Install-Module PnP.PowerShell -Scope CurrentUser
 *   Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SPGo-ModernAuth" -Tenant <yourtenant>.onmicrosoft.com
 *
 * siteUrl defaults to the Petrobras camap-agp/pci site. tenantId is
 * auto-resolved from the site if omitted (device code flow needs a
 * tenant-specific authority - 'organizations'/'common' fail with AADSTS50059).
 */

const CLIENT_ID = process.argv[2];
const SITE_URL = process.argv[3] || 'https://petrobrasbr.sharepoint.com/teams/camap-agp/pci';
const TENANT_ID_ARG = process.argv[4];

if (!CLIENT_ID) {
    console.error('Uso: node entra-access-probe.js <clientId> [siteUrl] [tenantId]');
    console.error('\nclientId é obrigatório - não existe mais um app compartilhado (o antigo "PnP Management Shell" foi apagado em 2024-09-09).');
    console.error('Crie o seu:');
    console.error('  Install-Module PnP.PowerShell -Scope CurrentUser');
    console.error('  Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SPGo-ModernAuth" -Tenant <seutenant>.onmicrosoft.com');
    process.exit(1);
}

const resource = new URL(SITE_URL).origin;
const scopes = [`${resource}/.default`];

function log(title, body) {
    console.log(`\n=== ${title} ===`);
    if (body !== undefined) console.log(body);
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function resolveTenantId() {
    if (TENANT_ID_ARG) return TENANT_ID_ARG;

    const res = await fetch(`${resource}/_vti_bin/client.svc`, { headers: { Authorization: 'Bearer' } });
    const wwwAuth = res.headers.get('www-authenticate') || '';
    const match = wwwAuth.match(/realm="([^"]+)"/);
    if (!match) {
        throw new Error(`Could not auto-resolve the Entra ID tenant for ${resource}. Pass it explicitly as the 4th argument.`);
    }
    return match[1];
}

function explainAadError(errBody) {
    const desc = (errBody && errBody.error_description) || (errBody && errBody.message) || JSON.stringify(errBody);
    const code = (String(desc).match(/AADSTS(\d+)/) || [])[1];
    const known = {
        '50059': 'Endpoint multi-tenant genérico não serve para device code -> bug corrigido neste script (agora resolve o tenant automaticamente).',
        '65001': 'Ninguém consentiu ainda com este app neste tenant, e consentimento de usuário parece estar desabilitado -> precisa de CONSENTIMENTO DE ADMINISTRADOR do tenant Petrobras para este clientId/escopo.',
        '50076': 'MFA é exigido e não foi satisfeito neste fluxo -> normal na primeira tentativa, responda ao desafio de MFA quando solicitado.',
        '53003': 'Bloqueado por política de CONDITIONAL ACCESS do tenant (ex: dispositivo não gerenciado/não compliant, localização, cliente não aprovado) -> pedir para o TI liberar este client/app ou fluxo para o seu usuário.',
        '70011': 'O escopo solicitado não é válido/reconhecido para este recurso -> revisar o escopo (AllSites.Write vs .default) com o TI.',
        '700016': 'Esse clientId não existe (ou não existe mais) na diretoria da Microsoft, ou não foi registrado/instalado neste tenant especificamente -> confira se o clientId está correto; se estiver, é preciso registrar um App Registration dedicado neste tenant (Register-PnPEntraIDAppForInteractiveLogin) e conceder consentimento de admin para ele.',
        '90094': 'Precisa de consentimento de administrador para as permissões solicitadas -> mesma causa do 65001/700016, pedir consentimento de admin.'
    };
    return code && known[code]
        ? `AADSTS${code}: ${known[code]}\n\nMensagem original: ${desc}`
        : desc;
}

// Speaks the device-code grant directly against the v2.0 endpoints:
// https://learn.microsoft.com/entra/identity-platform/v2-oauth2-device-code
async function acquireTokenByDeviceCode(tenantId) {
    const deviceCodeRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ client_id: CLIENT_ID, scope: scopes.join(' ') })
    });
    const deviceCodeBody = await deviceCodeRes.json();
    if (!deviceCodeRes.ok) {
        throw deviceCodeBody;
    }

    log('PASSO 1: login', deviceCodeBody.message);

    let intervalMs = (deviceCodeBody.interval || 5) * 1000;
    const expiresAt = Date.now() + (deviceCodeBody.expires_in || 900) * 1000;

    while (Date.now() < expiresAt) {
        await sleep(intervalMs);

        const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
                client_id: CLIENT_ID,
                device_code: deviceCodeBody.device_code
            })
        });
        const tokenBody = await tokenRes.json();

        if (tokenRes.ok) {
            return tokenBody;
        }
        if (tokenBody.error === 'authorization_pending') {
            continue;
        }
        if (tokenBody.error === 'slow_down') {
            intervalMs += 5000;
            continue;
        }
        throw tokenBody;
    }

    throw new Error('Device code expired before login was completed.');
}

async function testRestCall(path, accessToken) {
    const res = await fetch(`${SITE_URL}${path}`, {
        headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: 'application/json;odata=verbose'
        }
    });
    const text = await res.text();
    let body;
    try { body = JSON.parse(text); } catch (e) { body = text; }
    return { ok: res.ok, status: res.status, body };
}

async function main() {
    const tenantId = await resolveTenantId();
    log('Alvo', `site: ${SITE_URL}\nclientId: ${CLIENT_ID}\ntenant: ${tenantId}\nscopes: ${scopes.join(', ')}`);

    let tokenResult;
    try {
        tokenResult = await acquireTokenByDeviceCode(tenantId);
    } catch (err) {
        log('FALHA na autenticação (Entra ID recusou o login)', explainAadError(err));
        log('Conclusão', 'Isso já é evidência concreta para levar ao TI: o tenant/app está bloqueando este tipo de login antes mesmo de chegar no SharePoint.');
        process.exitCode = 1;
        return;
    }

    log('PASSO 1: login', `OK - token obtido (expira em ${tokenResult.expires_in}s)`);

    try {
        const title = await testRestCall("/_api/web/title", tokenResult.access_token);
        if (title.ok) {
            log('PASSO 2: leitura do site (GetWeb/title)', `OK - título do site: "${title.body.d.Title}"`);
        } else {
            log('PASSO 2: leitura do site FALHOU', `HTTP ${title.status}: ${JSON.stringify(title.body)}`);
            log('Conclusão', 'Login no Entra ID funcionou, mas o SharePoint recusou o acesso a este site -> provavelmente falta permissão (Sites.Selected) ou consentimento de admin específico para SharePoint.');
            process.exitCode = 1;
            return;
        }

        const user = await testRestCall("/_api/web/currentuser", tokenResult.access_token);
        if (user.ok) {
            log('PASSO 3: identidade no site (currentuser)', `OK - logado como: ${user.body.d.Title} (${user.body.d.Email})`);
        }

        log('RESULTADO', 'Autenticação moderna (MFA) + acesso de leitura ao site FUNCIONAM. Pode prosseguir para testar escrita/publicação pela extensão.');
    } catch (err) {
        log('Erro inesperado ao chamar o SharePoint', err.message || err);
        process.exitCode = 1;
    }
}

main();
