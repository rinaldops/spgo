'use strict';

import * as vscode from 'vscode';
import { spfi, SPFI } from '@pnp/sp';
import '@pnp/sp/presets/all.js';
import { SPDefault } from '@pnp/nodejs';
import { TimelinePipe } from '@pnp/core';
import { Queryable } from '@pnp/queryable';

import { IAppManager, IConfig } from '../spgo';
import { ModernAuthService } from './modernAuthService';

// @pnp/nodejs ships a ready-made MSAL behavior, but it only does app-only auth
// (ConfidentialClientApplication.acquireTokenByClientCredential) - no user in the loop,
// so no MFA. This behavior plugs PnPjs's own auth pipeline into the device-code login
// already implemented (and validated against a real tenant) in ModernAuthService,
// instead of duplicating token acquisition/caching logic here.
function MSALDeviceCode(config: IConfig): TimelinePipe<Queryable> {
    return (instance: Queryable) => {
        // init typed as `any` - RequestInit isn't declared without a DOM/newer-@types/node
        // lib, same gap the ambient fetch/URL declarations in src/types/global.d.ts work around.
        instance.on.auth(async (url: URL, init: any) => {
            let appManager: IAppManager = vscode.window.spgo;
            await ModernAuthService.acquireToken(appManager, config);
            init.headers = Object.assign({}, init.headers, { Authorization: `Bearer ${appManager.credentials.accessToken}` });
            return [url, init];
        });
        return instance;
    };
}

const spCache: Map<string, SPFI> = new Map();

// Configured @pnp/sp instance for a given site, using Modern (device-code/MFA) auth.
// @pnp/sp is what actually replaces spsave/sppull for the Modern-auth path - it already
// implements chunked upload (files.addChunked) and folder recursion; SPGo only needs to
// supply the auth behavior above.
export class PnpService {
    static getSp(siteUrl: string, config: IConfig): SPFI {
        if (!spCache.has(siteUrl)) {
            spCache.set(siteUrl, spfi(siteUrl).using(SPDefault({ baseUrl: siteUrl }), MSALDeviceCode(config)));
        }
        return spCache.get(siteUrl);
    }
}
