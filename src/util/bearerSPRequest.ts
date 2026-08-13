'use strict';

export interface IBearerRequestOptions {
    body? : any;
    rawBody? : Buffer;
    headers? : any;
}

export interface IBearerResponse {
    body : any;
    statusCode : number;
}

// Minimal drop-in replacement for the subset of sp-request's ISPRequest interface
// that SPGo actually calls (requestDigest/get/post), built on native fetch + a
// pre-acquired OAuth Bearer token. node-sp-auth (which sp-request wraps) only
// understands unattended credential strategies (Digest/NTLM/Forms/ADFS/AddinOnly)
// and has no modern-auth/MFA strategy, so it can't be reused here.
export class BearerSPRequest {

    constructor(private accessToken : string) {}

    requestDigest(siteUrl : string) : Promise<string> {
        return this.post(`${siteUrl}/_api/contextinfo`, { body: {} })
            .then(response => response.body.d.GetContextWebInformation.FormDigestValue);
    }

    get(url : string, options : IBearerRequestOptions = {}) : Promise<IBearerResponse> {
        return this.send('GET', url, options);
    }

    post(url : string, options : IBearerRequestOptions = {}) : Promise<IBearerResponse> {
        return this.send('POST', url, options);
    }

    // Raw binary download (e.g. GetFileByServerRelativeUrl('...')/$value) - response isn't JSON.
    getBinary(url : string) : Promise<Buffer> {
        return fetch(url, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${this.accessToken}` }
        })
        .then((res : any) => {
            if (!res.ok) {
                throw new Error(`SharePoint request failed (${res.status} ${res.statusText}): ${url}`);
            }
            return res.arrayBuffer();
        })
        .then((buf : ArrayBuffer) => Buffer.from(buf));
    }

    private send(method : string, url : string, options : IBearerRequestOptions) : Promise<IBearerResponse> {
        const headers : any = Object.assign({
            'Authorization': `Bearer ${this.accessToken}`,
            'Accept': 'application/json;odata=verbose'
        }, options.headers || {});

        let body : any;
        if (options.rawBody) {
            body = options.rawBody;
        } else if (options.body && Object.keys(options.body).length) {
            headers['Content-Type'] = 'application/json;odata=verbose';
            body = JSON.stringify(options.body);
        }

        return fetch(url, { method, headers, body })
            .then((res : any) => res.text().then((text : string) => ({ res, text })))
            .then(({ res, text } : any) => {
                let parsedBody : any = {};
                if (text) {
                    try { parsedBody = JSON.parse(text); } catch (e) { parsedBody = text; }
                }
                if (!res.ok) {
                    const message = (parsedBody && parsedBody.error && parsedBody.error.message && parsedBody.error.message.value) || res.statusText;
                    throw new Error(`SharePoint request failed (${res.status}): ${message}`);
                }
                return { body: parsedBody, statusCode: res.status };
            });
    }
}
