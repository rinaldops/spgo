'use strict';

import * as vscode from 'vscode';
import { PublicClientApplication, AccountInfo, AuthenticationResult, Configuration } from '@azure/msal-node';

import { IAppManager, IConfig } from '../spgo';
import { Logger } from '../util/logger';

// There is no shared/well-known public client to fall back to: the community "PnP
// Management Shell" app that used to fill this role was deleted by the PnP team on
// 2024-09-09 for security reasons. Every tenant now needs its own App Registration -
// set authenticationDetails.clientId in SPGo.json. The fastest way to create one:
//   Install-Module PnP.PowerShell -Scope CurrentUser
//   Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SPGo-ModernAuth" -Tenant <yourtenant>.onmicrosoft.com

const pcaCache : Map<string, PublicClientApplication> = new Map();

export class ModernAuthService {

    // Acquires (or reuses/silently refreshes) an OAuth access token for the site's
    // SharePoint resource and stores it on appManager.credentials for this session.
    static acquireToken(appManager : IAppManager, config : IConfig) : Promise<IAppManager> {

        const details : any = config.authenticationDetails || {};
        const clientId : string = details.clientId;
        const resource : string = new URL(config.sharePointSiteUrl).origin;
        const scopes : string[] = [`${resource}/.default`];

        if (!clientId) {
            return Promise.reject(new Error(
                'Modern authentication needs authenticationDetails.clientId in SPGo.json - there is no shared app to fall back to anymore ' +
                '(the community "PnP Management Shell" app was deleted 2024-09-09). Register your own with: ' +
                'Install-Module PnP.PowerShell -Scope CurrentUser; ' +
                'Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SPGo-ModernAuth" -Tenant <yourtenant>.onmicrosoft.com'
            ));
        }

        return ModernAuthService.resolveTenantId(details.tenantId, resource)
            .then((tenantId : string) => {
                const pca : PublicClientApplication = ModernAuthService.getPca(clientId, tenantId);

                return ModernAuthService.acquireSilent(pca, appManager, scopes)
                    .catch(() => ModernAuthService.acquireByDeviceCode(pca, scopes));
            })
            .then((result : AuthenticationResult) => {
                appManager.credentials.accessToken = result.accessToken;
                appManager.credentials.accessTokenExpiresOn = result.expiresOn;
                appManager.credentials.msalAccount = result.account;
                return appManager;
            });
    }

    // Device code flow requires a tenant-specific authority - the multi-tenant
    // 'organizations'/'common' endpoints fail with AADSTS50059. If authenticationDetails
    // doesn't pin a tenantId, auto-resolve it from the site itself via SharePoint's
    // unauthenticated realm-discovery response (WWW-Authenticate: Bearer realm="<tenantId>"),
    // the same trick PnP/SPO tooling uses - no need to know/guess the tenant GUID or
    // verified onmicrosoft.com domain.
    private static resolveTenantId(configuredTenantId : string, resource : string) : Promise<string> {
        if (configuredTenantId) {
            return Promise.resolve(configuredTenantId);
        }

        return fetch(`${resource}/_vti_bin/client.svc`, { headers: { 'Authorization': 'Bearer' } })
            .then((res : any) => {
                const wwwAuth : string = res.headers.get('www-authenticate') || '';
                const match = wwwAuth.match(/realm="([^"]+)"/);
                if (!match) {
                    throw new Error(`Could not auto-resolve the Entra ID tenant for ${resource}. Set authenticationDetails.tenantId in SPGo.json explicitly.`);
                }
                return match[1];
            });
    }

    private static getPca(clientId : string, tenantId : string) : PublicClientApplication {
        const key : string = `${clientId}|${tenantId}`;
        if (!pcaCache.has(key)) {
            const msalConfig : Configuration = {
                auth: {
                    clientId: clientId,
                    authority: `https://login.microsoftonline.com/${tenantId}`
                }
            };
            pcaCache.set(key, new PublicClientApplication(msalConfig));
        }
        return pcaCache.get(key);
    }

    private static acquireSilent(pca : PublicClientApplication, appManager : IAppManager, scopes : string[]) : Promise<AuthenticationResult> {
        const account : AccountInfo = appManager.credentials && appManager.credentials.msalAccount;
        if (!account) {
            return Promise.reject('no cached account');
        }
        return pca.acquireTokenSilent({ account, scopes });
    }

    private static acquireByDeviceCode(pca : PublicClientApplication, scopes : string[]) : Promise<AuthenticationResult> {
        return pca.acquireTokenByDeviceCode({
            scopes,
            deviceCodeCallback: (response) => {
                Logger.outputMessage(response.message, vscode.window.spgo.outputChannel);
                vscode.window.showInformationMessage(
                    `SPGo: entre com o código ${response.userCode} em ${response.verificationUri} para autenticar no SharePoint.`,
                    'Abrir navegador'
                ).then(choice => {
                    if (choice === 'Abrir navegador') {
                        vscode.env.openExternal(vscode.Uri.parse(response.verificationUri));
                    }
                });
            }
        });
    }
}
