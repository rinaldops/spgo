'use strict';

// import * as vscode from 'vscode';

import * as fse from 'fs-extra';
import * as path from 'path';
import { SPFI } from '@pnp/sp';
import * as globToRegExp from 'glob-to-regexp';
import { Uri } from 'vscode';
import { RequestHelper } from '../util/requestHelper';
import { ISPRequest, IAuthOptions } from 'sp-request';
import { SPPull, ISPPullContext, ISPPullOptions } from 'sppull';
import { WorkspaceHelper } from '../util/workspaceHelper';
import { spsave, ICoreOptions, FileOptions } from 'spsave';
import { PnpService } from '../service/pnpService';
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
    // OAuth/MFA strategy (it wraps node-sp-auth), so this goes through @pnp/sp instead,
    // configured with Modern Auth's device-code behavior (see PnpService).
    public uploadFileModern(fileUri : Uri, fileBuffer : Buffer, checkin : boolean, checkinType : number, checkinMessage : string) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(fileUri.toString(), this._config);
        let folderPath : string = fileUri.path.substring(0, fileUri.path.lastIndexOf('/'));
        let fileName : string = fileUri.path.substring(fileUri.path.lastIndexOf('/') + 1);
        let sp : SPFI = PnpService.getSp(sharePointSiteUrl.toString(), this._config);

        return sp.web.getFolderByServerRelativePath(folderPath).files.addChunked(fileName, fileBuffer, { Overwrite: true })
            .then(fileInfo => checkin
                ? sp.web.getFileByServerRelativePath(fileInfo.ServerRelativeUrl).checkin(checkinMessage || '', checkinType).then(() => fileInfo)
                : fileInfo);
    }

    // Modern Authentication equivalent of downloadFiles() for a single file - sppull has no
    // OAuth/MFA strategy either, so this goes through @pnp/sp instead.
    public downloadFileModern(fileUri : Uri, localDestPath : string) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(fileUri.toString(), this._config);
        let sp : SPFI = PnpService.getSp(sharePointSiteUrl.toString(), this._config);

        return sp.web.getFileByServerRelativePath(fileUri.path).getBuffer()
            .then(buffer => {
                fse.ensureDirSync(path.dirname(localDestPath));
                fse.writeFileSync(localDestPath, Buffer.from(buffer));
                return [{ SavedToLocalPath: localDestPath }];
            });
    }

    // Modern Authentication equivalent of publishWorkspace()'s bulk upload - walks the local
    // folder, filters against the configured glob patterns (same glob-to-regexp already used
    // by DownloadFileOptionsFactory), and uploads each match via @pnp/sp's chunked upload.
    // globPatterns is optional - pass null/empty to upload everything under localRoot
    // unfiltered (right-click "publish" on an arbitrary folder: send it all, recursively).
    public uploadFolderModern(localRoot : string, remoteFolderUri : Uri, globPatterns : string[], checkin : boolean, checkinType : number, checkinMessage : string) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(remoteFolderUri.toString(), this._config);
        let sp : SPFI = PnpService.getSp(sharePointSiteUrl.toString(), this._config);
        let matchers : RegExp[] = (globPatterns && globPatterns.length)
            ? globPatterns.map(pattern => globToRegExp(pattern, { flags: 'i', globstar: true }))
            : null;

        let localFiles : string[] = this.walkLocalFiles(localRoot)
            .filter(filePath => !matchers || matchers.some(re => re.test(path.relative(localRoot, filePath).replace(/\\/g, '/'))));

        return Promise.all(localFiles.map(filePath => {
            let relative : string = path.relative(localRoot, filePath).replace(/\\/g, '/');
            let remoteFolder : string = (remoteFolderUri.path + '/' + path.dirname(relative)).replace(/\/\.$/, '').replace(/\/+/g, '/');
            let fileBuffer : Buffer = fse.readFileSync(filePath);

            return sp.web.getFolderByServerRelativePath(remoteFolder).files.addChunked(path.basename(filePath), fileBuffer, { Overwrite: true })
                .then(fileInfo => checkin
                    ? sp.web.getFileByServerRelativePath(fileInfo.ServerRelativeUrl).checkin(checkinMessage || '', checkinType).then(() => fileInfo)
                    : fileInfo);
        }));
    }

    // Modern Authentication equivalent of downloadFiles()'s bulk download (populateWorkspace /
    // retrieveFolder) - recursively mirrors a SharePoint folder to disk via @pnp/sp.
    public downloadFolderModern(remoteFolderUri : Uri, localDestRoot : string) : Promise<any>{

        let sharePointSiteUrl : Uri = WorkspaceHelper.getSiteUriForActiveWorkspace(remoteFolderUri.toString(), this._config);
        let sp : SPFI = PnpService.getSp(sharePointSiteUrl.toString(), this._config);

        return this.walkRemoteFolder(sp, remoteFolderUri.path, remoteFolderUri.path, localDestRoot);
    }

    // sppull's bulk download also skips each library's system "Forms" folder - this does too.
    private async walkRemoteFolder(sp : SPFI, remotePath : string, basePath : string, localDestRoot : string) : Promise<any[]> {

        let folder = sp.web.getFolderByServerRelativePath(remotePath);
        let [files, subFolders] = await Promise.all([folder.files(), folder.folders()]);
        let results : any[] = [];

        for (let file of files) {
            let buffer : ArrayBuffer = await sp.web.getFileByServerRelativePath(file.ServerRelativeUrl).getBuffer();
            let localPath : string = path.join(localDestRoot, file.ServerRelativeUrl.substring(basePath.length));

            fse.ensureDirSync(path.dirname(localPath));
            fse.writeFileSync(localPath, Buffer.from(buffer));
            results.push({ SavedToLocalPath: localPath });
        }

        for (let subFolder of subFolders) {
            if (subFolder.Name === 'Forms') { continue; }
            results = results.concat(await this.walkRemoteFolder(sp, subFolder.ServerRelativeUrl, basePath, localDestRoot));
        }

        return results;
    }

    // fs-extra has no recursive walk built in, and we only need a flat list of absolute
    // file paths - not worth a dependency like klaw for that.
    // (withFileTypes support predates this project's @types/node - cast around the gap,
    // same as the ambient fetch/URL declarations in src/types/global.d.ts.)
    private walkLocalFiles(dir : string) : string[] {
        let results : string[] = [];
        let entries : { name : string, isDirectory() : boolean }[] = fse.readdirSync(dir, { withFileTypes: true }) as any;
        for (let entry of entries) {
            let entryPath : string = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                results = results.concat(this.walkLocalFiles(entryPath));
            } else {
                results.push(entryPath);
            }
        }
        return results;
    }
}
