'use strict';

// import * as vscode from 'vscode';

import * as fse from 'fs-extra';
import * as path from 'path';
import { Uri } from 'vscode';
import { RequestHelper } from '../util/requestHelper';
import { BearerSPRequest } from '../util/bearerSPRequest';
import { ISPRequest, IAuthOptions } from 'sp-request';
import { SPPull, ISPPullContext, ISPPullOptions } from 'sppull';
import { WorkspaceHelper } from '../util/workspaceHelper';
import { spsave, ICoreOptions, FileOptions } from 'spsave';
import { ISPFileInformation, IConfig, IFileGateway } from '../spgo';

export class SPFileGateway implements IFileGateway{

    _config : IConfig;

    constructor(config : IConfig){
        this._config = config;
    }

    public checkOutFile(fileUri : Uri, spr : ISPRequest ) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(fileUri.toString(), this._config);

        return spr.requestDigest(sharePointSiteUrl.toString())
            .then(digest => {
                return spr.post(sharePointSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + encodeURI(fileUri.path) +"')/CheckOut()", {
                    body: {},
                    headers: RequestHelper.createAuthHeaders(this._config, digest)
                });
            });
    }

    public deleteFile(fileUri : Uri, spr : ISPRequest ) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(this._config.sharePointSiteUrl + fileUri, this._config);

        return spr.requestDigest(sharePointSiteUrl.toString())
            .then(digest => {
                return spr.post(sharePointSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + encodeURI(fileUri.path) +"')", {
                    body: {},
                    headers: RequestHelper.createAuthHeaders(this._config, digest, {
                        'X-HTTP-Method':'DELETE',
                        'accept': 'application/json; odata=verbose',
                        'content-type': 'application/json; odata=verbose'
                    })
                });
            });
    }

    public downloadFiles(context : ISPPullContext, fileOptions : ISPPullOptions) : Promise<any>{
        return SPPull.download(context, fileOptions);
    }

    // CheckOutType: Online = 0; Offline = 1; None = 2.
    // all status values: https://msdn.microsoft.com/en-us/library/office/dn450841.aspx
    public getFileInformation( fileUri : Uri, spr : ISPRequest ) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(this._config.sharePointSiteUrl + fileUri, this._config);

        return spr.requestDigest(sharePointSiteUrl.toString())
            .then(digest => {
                return spr.get(sharePointSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + encodeURI(fileUri.path) +"')/?$select=Name,ServerRelativeUrl,CheckOutType,TimeLastModified,CheckedOutByUser", {
                    body: {},
                    headers: RequestHelper.createAuthHeaders(this._config, digest)
                })
                .then( response => {
                    let fileInfo : ISPFileInformation = {
                        checkOutType : response.body.d.CheckOutType,
                        name : response.body.d.Name,
                        timeLastModified : response.body.d.TimeLastModified
                    }

                    // File is checked out
                    if( fileInfo.checkOutType == 0 || fileInfo.checkOutType == 1){
                        // '/_api/web/getFileByServerRelativeUrl(\'' + encodeURI(fileName) + '\')/CheckedOutByUser?$select=Title,Email';
                        return spr.get(sharePointSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + encodeURI(fileUri.path) +"')/CheckedOutByUser?$select=Title,Email", {
                                body: {},
                                headers: RequestHelper.createAuthHeaders(this._config, digest)
                            }).then( userInfo => {
                                fileInfo.checkOutBy = userInfo.body.d.Title;
                                return fileInfo;
                            });
                    }
                    else{
                        return fileInfo;
                    }
                })
            })
    }

    public undoCheckOutFile(fileUri : Uri, spr : ISPRequest ) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(this._config.sharePointSiteUrl + fileUri, this._config);

        return spr.requestDigest(sharePointSiteUrl.toString())
            .then(digest => {
                return spr.post(sharePointSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + encodeURI(fileUri.path) +"')/undocheckout()", {
                    body: {},
                    headers: RequestHelper.createAuthHeaders(this._config, digest)
                });
            });
    }

    public uploadFiles(coreOptions : ICoreOptions, credentials : IAuthOptions, fileOptions : FileOptions) : Promise<any>{

        RequestHelper.setNtlmHeader(this._config)

        return spsave(coreOptions, credentials, fileOptions);
    }

    // Modern Authentication equivalent of uploadFiles() for a single file - spsave has no
    // OAuth/MFA strategy (it wraps node-sp-auth), so this bypasses it and calls the
    // Files/add REST endpoint directly with the Bearer token instead.
    public uploadFileModern(fileUri : Uri, fileBuffer : Buffer, spr : ISPRequest, checkin : boolean, checkinType : number, checkinMessage : string) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(fileUri.toString(), this._config);
        let folderPath : string = fileUri.path.substring(0, fileUri.path.lastIndexOf('/'));
        let fileName : string = fileUri.path.substring(fileUri.path.lastIndexOf('/') + 1);

        return spr.requestDigest(sharePointSiteUrl.toString())
            .then(digest => {
                return (spr as any).post(sharePointSiteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURI(folderPath) + "')/Files/add(url='" + encodeURIComponent(fileName) + "',overwrite=true)", {
                    rawBody: fileBuffer,
                    headers: RequestHelper.createAuthHeaders(this._config, digest)
                }).then(() => {
                    if (checkin) {
                        let fileServerRelativeUrl : string = folderPath + '/' + fileName;
                        return spr.post(sharePointSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + encodeURI(fileServerRelativeUrl) + "')/CheckIn(comment='" + encodeURIComponent(checkinMessage || '') + "',checkintype=" + checkinType + ")", {
                            body: {},
                            headers: RequestHelper.createAuthHeaders(this._config, digest)
                        });
                    }
                });
            });
    }

    // Modern Authentication equivalent of downloadFiles() for a single file - sppull has no
    // OAuth/MFA strategy either, so this fetches the raw file content directly and writes
    // it to localDestPath, mirroring what SPPull.download() would have produced.
    public downloadFileModern(fileUri : Uri, spr : ISPRequest, localDestPath : string) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(fileUri.toString(), this._config);
        let bearer : BearerSPRequest = spr as any;

        return bearer.getBinary(sharePointSiteUrl + "/_api/web/GetFileByServerRelativeUrl('" + encodeURI(fileUri.path) + "')/$value")
            .then(buffer => {
                fse.ensureDirSync(path.dirname(localDestPath));
                fse.writeFileSync(localDestPath, buffer);
                return [{ SavedToLocalPath: localDestPath }];
            });
    }
}
