import { Injectable, Self } from '@angular/core';
import { IKL } from "../../Attach2Dynamics.CrmJs"
import { Subject, Subscription } from 'rxjs';
import { A2dAppService } from '../a2d-app.service';
import { GridData } from '../shared/grid-definition';
import { DocumentLocation } from '../shared/document-location';
import { Connector } from '../shared/connector';
import { EntityConfiguration } from '../shared/entity-configuration';
import { ModalService } from '../shared/modal/modal.service';
import { UtilityService } from '../shared/utility/utility.service';
import { SpinnerService } from '../core/spinner/spinner.service';
import { BreadcrumbService } from '../connector-tab/breadcrumb/breadcrumb.service';
import { HttpClient, HttpEvent, HttpEventType, HttpHeaders, HttpResponse } from '@angular/common/http';
import { subtract } from 'ngx-bootstrap/chronos';
import { Console, debug, error } from 'console';
import { GridService } from '@app/connector-tab/grid/grid.service';
import { Field, CRMEntity } from '@app/shared/metadata-json';
import { strict } from 'assert';
import { Header } from 'primeng/api';
import { lookup } from 'dns';
import { Guid } from 'ikl_/Scripts/xrm-webapi';
import { subscribe } from 'diagnostics_channel';
import { EntityAndRoles, SecurityPrivilege, UserPrivilegeData } from '@app/shared/EntityAndRoles';
//import { EntityAndRoles, SecurityPrivilege } from '@app/shared/EntityAndRoles';
declare let download: any;
declare let crmWebApi: any;

@Injectable({
    providedIn: 'root'
})
export class SharepointService {
    actionName: string = "ikl_SharePointCore";
    retrieveDocumentLocationsSub = new Subscription();
    getFilesSub = new Subscription();
    createFoldersSub = new Subscription();
    uploadFileSub = new Subscription();
    downloadSub = new Subscription();
    renameFileSub = new Subscription();
    viewFileSub = new Subscription();
    renameFolderSub = new Subscription();
    shareLinkSub = new Subscription();
    searchSub = new Subscription();
    actionOutputSub = new Subscription();
    fillGridSub = new Subscription();
    deleteFileSub = new Subscription();
    createFolderAndUploadFilesSub = new Subscription();
    uploadFilesStartSub = new Subscription();
    uploadFileSPSub = new Subscription();
    downloadSPSub = new Subscription();
    retrieveDocumentLocations$ = new Subject<any>();
    getFiles$ = new Subject<any>();
    createFolders$ = new Subject<any>();
    uploadFile$ = new Subject<any>();
    viewFile$ = new Subject<any>();
    download$ = new Subject<any>();
    shareLink$ = new Subject<any>();
    search$ = new Subject<any>();
    App_AddLookUpListItems$ = new Subject<any>();
    App_GetListName$ = new Subject<any>();
    App_AddItemToList$ = new Subject<any>();
    App_AddUserListItems$ = new Subject<any>();
    _Xrm: IKL.Attach2Dynamics.CrmJs; //Shreyas: 8 April 2022

    createSharePointMetadataSub = new Subscription();
    createSharePointMetadata$ = new Subject<any>();
    getFilesAngularSub = new Subscription();
    getFilesAngular$ = new Subject<any>();
    createViewXmlSub = new Subscription();
    createViewXml$ = new Subject<any>();
    retrieveRecord$ = new Subject<any>();
    retrieveRecordSub = new Subscription();
    getMainRootLibraryViewsSub = new Subscription();
    App_AddLookUpListItemsSub = new Subscription();
    App_GetListNameSub = new Subscription();
    App_AddItemToListSub = new Subscription();
    App_AddUserListItemsSub = new Subscription();
    getMainRootLibraryViews$ = new Subject<any>();
    getSharePointLookUpValues$ = new Subject<any>();
    /********** Messages **********/
    message_wronglyConfiguredEC = this.a2dAppService.labelsMultiLanguage['entityconfigerror'];

    /********** End Messages **********/

    // #Added 23/09/2019
    validateFolderInSharePointSub = new Subscription();

    //shrujan 13 feb 23 for D&D
    moveFile$ = new Subject<any>();
    moveFileSPSub = new Subscription();

    downloadSPFileCount: number = 0;
    viewXmlArray = [];
    colFields = [];
    acceessToken: any;
    currentPage: number = 1;
    pageSize: number = 50;
    nextRowDataHref: string;
    checkPasswordChengedSub = new Subscription();
    //Shrujan
    generateSPAccessToken$ = new Subject<any>();
    generateSPAccessTokenSub = new Subscription();

    syncStatues$ = new Subject<any>();
    syncStatuesSub = new Subscription();
    spdocloc$ = new Subject<any>();
    spdoclocSub = new Subscription();
    entitySecurityMetadataFetch$ = new Subject<any>();
    entitySecurityMetadataFetchSub = new Subscription();

    userMaxPrev$ = new Subject<any>();
    userMaxPrevSub = new Subscription();

    permissionId$ = new Subject<any>();
    permissionIdSub = new Subscription();

    sss_GroupCollection$ = new Subject<any>();
    sss_GroupCollectionSub = new Subscription();

    isIkl_FilePrivillagesValid: boolean = true;
    folderId: string;

    constructor(private http: HttpClient, private a2dAppService: A2dAppService, private modalService: ModalService,
        private utilityService: UtilityService, private spinnerService: SpinnerService,
        private breadcrumbService: BreadcrumbService, private gridService: GridService) { }


    //#region New upload method
    /** 
     * New upload method to upload large files (1.5 GB)  in chunks.
     * @param file 
     * @param decryptedToken 
     * @param folderPath 
     * @param fileName 
     * @param selectedConnectorTab 
     * @param selectedEntityConfiguration 
     * @returns 
     */
    uploadFilesSPmain(file: any, folderPath: string, fileName: string, selectedConnectorTab: any, selectedEntityConfiguration: any, source?: string) {
        return new Promise((resolve, reject) => {
            let functionName: string = "uploadFilesSPmain";
            let self = this;
            let folderRelativePath: string;
            let firstCharOfPath: string;
            let serverRelativeURL: string;
            let encodedFileName: string;
            let fileUniqueId: string;
            let itemID: string;
            folderRelativePath = folderPath.replace(/'/g, "''");
            fileName = fileName.replace(/'/g, "''");
            firstCharOfPath = folderPath.charAt(0);

            //@ts-ignore
            this.acceessToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(selectedConnectorTab.access_token).DecryptedValue;

            //Checking if path have "/"" character at first position if contains then remove because it is not supports inrequest url (GetFolderByServerRelativeUrl).
            if (firstCharOfPath == "/") {
                serverRelativeURL = encodeURIComponent(folderRelativePath.substring(1, folderRelativePath.length));
            }
            else {
                serverRelativeURL = encodeURIComponent(folderRelativePath);
            }

            // Calling create file and then devide file into chunks and upload through rest pi's
            this.createFileFirst(file, this.acceessToken, fileName, folderPath, serverRelativeURL, selectedConnectorTab, selectedEntityConfiguration, source).then((result: any) => {
                if (result.hasError == true) {
                    if ('comments' in result) {
                        if (result.comments.includes("tokenExpired")) {
                            this.createFileFirst(file, result.newToken, fileName, folderPath, serverRelativeURL, selectedConnectorTab, selectedEntityConfiguration, source).then((result: any) => {
                                fileUniqueId = result.UniqueId;
                                itemID = result.ListItemAllFields.ID;
                                fileName = result.Name;
                                fileName = fileName.replace(/'/g, "''");
                                //encodedFileName = encodeURIComponent(fileName);
                                let fr = new FileReader();
                                let offset = 0;
                                let sharePointSite: string = "";
                                let fileServerRelativeURL: string = "";
                                // the total file size in bytes...    
                                sharePointSite = selectedConnectorTab.absolute_url;
                                if (sharePointSite.split('/').length > 3) {
                                    let sitePath: any = sharePointSite.split('/').splice(3).join('/');
                                    fileServerRelativeURL = '/' + sitePath + "/" + serverRelativeURL;
                                }
                                else {
                                    fileServerRelativeURL = "/" + serverRelativeURL;
                                }
                                let total = file.size;
                                // 5MB Chunks as represented in bytes (if the file is less than a MB, seperate it into two chunks of 80% and 20% the size)...    
                                let length = 5000000 > total ? Math.round(total * 0.8) : 5000000;
                                let chunks = [];
                                //reads in the file using the fileReader HTML5 API (as an ArrayBuffer) - readAsBinaryString is not available in IE!    
                                fr.readAsArrayBuffer(file);
                                fr.onload = (evt: any) => {
                                    while (offset < total) {
                                        //if we are dealing with the final chunk, we need to know...    
                                        if (offset + length > total) {
                                            length = total - offset;
                                        }
                                        //work out the chunks that need to be processed and the associated REST method (start, continue or finish)    
                                        chunks.push({
                                            offset,
                                            length,
                                            method: this.getUploadMethod(offset, length, total)
                                        });
                                        offset += length;
                                    }
                                    //each chunk is worth a percentage of the total size of the file...    
                                    const chunkPercentage = (total / chunks.length) / total * 100;
                                    if (chunks.length > 0) {
                                        //the unique guid identifier to be used throughout the upload session    
                                        const id = this.generateGUID();
                                        //Start the upload - send the data to S    
                                        this.uploadFileStart(evt.target.result, id, fileServerRelativeURL, folderPath, fileName, itemID, fileUniqueId, chunks, 0, 0, chunkPercentage, resolve, reject, selectedConnectorTab, selectedEntityConfiguration, this.acceessToken);
                                    }
                                };

                            }).catch(err => {
                                reject(err);
                            });
                        }
                        else {
                            let response = {};
                            response["FileName"] = fileName;
                            response["FilePath"] = serverRelativeURL;
                            response["status"] = "false";
                            response["FileUniqueId"] = fileUniqueId;
                            resolve(response);
                        }
                    }
                    else {
                        let response = {};
                        response["FileName"] = fileName;
                        response["FilePath"] = serverRelativeURL;
                        response["status"] = "false";
                        response["FileUniqueId"] = fileUniqueId;
                        resolve(response);
                    }
                }
                else {
                    fileUniqueId = result.UniqueId;
                    itemID = result.ListItemAllFields.ID;
                    fileName = result.Name;
                    fileName = fileName.replace(/'/g, "''");
                    encodedFileName = encodeURIComponent(fileName);
                    let fr = new FileReader();
                    let offset = 0;
                    let sharePointSite: string = "";
                    let fileServerRelativeURL: string = "";
                    // the total file size in bytes...    
                    sharePointSite = selectedConnectorTab.absolute_url;
                    if (sharePointSite.split('/').length > 3) {
                        let sitePath: any = selectedConnectorTab.absolute_url.split("/").splice(3).join('/');
                        sitePath = "/" + sitePath;
                        fileServerRelativeURL = sitePath + "/" + serverRelativeURL;
                    }
                    else {
                        fileServerRelativeURL = "/" + serverRelativeURL;
                    }
                    let total = file.size;
                    // 5MB Chunks as represented in bytes (if the file is less than a MB, seperate it into two chunks of 80% and 20% the size)...    
                    let length = 5000000 > total ? Math.round(total * 0.8) : 5000000;
                    let chunks = [];
                    //reads in the file using the fileReader HTML5 API (as an ArrayBuffer) - readAsBinaryString is not available in IE!    
                    fr.readAsArrayBuffer(file);
                    fr.onload = (evt: any) => {
                        while (offset < total) {
                            //if we are dealing with the final chunk, we need to know...    
                            if (offset + length > total) {
                                length = total - offset;
                            }
                            //work out the chunks that need to be processed and the associated REST method (start, continue or finish)    
                            chunks.push({
                                offset,
                                length,
                                method: this.getUploadMethod(offset, length, total)
                            });
                            offset += length;
                        }
                        //each chunk is worth a percentage of the total size of the file...    
                        const chunkPercentage = (total / chunks.length) / total * 100;
                        if (chunks.length > 0) {
                            //the unique guid identifier to be used throughout the upload session    
                            const id = this.generateGUID();
                            //Start the upload - send the data to S    
                            this.uploadFileStart(evt.target.result, id, fileServerRelativeURL, folderPath, fileName, itemID, fileUniqueId, chunks, 0, 0, chunkPercentage, resolve, reject, selectedConnectorTab, selectedEntityConfiguration, this.acceessToken);
                        }
                    };
                }
            }).catch(err => {
                let response = {};
                response["FileName"] = fileName;
                response["FilePath"] = serverRelativeURL;
                response["status"] = "false";
                response["FileUniqueId"] = fileUniqueId;
                if (!self.a2dAppService.isValid(err.code)) {
                    resolve(response);
                }
                else {
                    self.spinnerService.hide();
                    self.modalService.UploadStatusModalRef.hide();
                    self.a2dAppService.logError('', JSON.stringify(err), selectedEntityConfiguration, '', fileName, serverRelativeURL);
                    self.modalService.isOpen = true;
                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                }
            });
        });
    }

    createFileFirst(file, decryptedToken, fileName, folderPath, serverRelativeURL, selectedConnectorTab, selectedEntityConfiguration, source) {
        let functionName: string = "createFileFirst";
        let self = this;
        return new Promise((resolve, reject) => {
            // Construct the endpoint - The GetList method is available for SharePoint Online only. 
            let httpOptions = {};
            let isOverride: boolean;
            let stringIsOverride: string;
            let sharePointSite: string;
            let requestUrl: string;
            let isVersionEnabled: string;
            let encodedFileName: string;

            isOverride = selectedEntityConfiguration.length > 0 ? selectedEntityConfiguration[0].isOverride : selectedEntityConfiguration.isOverride;

            stringIsOverride = isOverride ? "true" : "false";

            isVersionEnabled = isOverride ? "false" : "true";

            sharePointSite = selectedConnectorTab.absolute_url;

            encodedFileName = encodeURIComponent(fileName);

            if ((this.a2dAppService.isValid(source) && source == "UploadFolder")) {
                //GetFolderByServerRelativePath(decodedurl='<server-relative-path>')
                requestUrl = sharePointSite + "/_api/web/GetFolderByServerRelativeUrl('" + serverRelativeURL + "')/Files/AddUsingPath(DecodedUrl=@a1, overwrite='" + stringIsOverride + "', EnsureUniqueFileName='" + isVersionEnabled + "')?@a1='" + encodedFileName + "'&$expand=ListItemAllFields"
            }
            else if (!this.a2dAppService.isValid(this.folderId)) {
                requestUrl = sharePointSite + "/_api/web/GetFolderByServerRelativeUrl('" + serverRelativeURL + "')/Files/AddUsingPath(DecodedUrl=@a1, overwrite='" + stringIsOverride + "', EnsureUniqueFileName='" + isVersionEnabled + "')?@a1='" + encodedFileName + "'&$expand=ListItemAllFields"
            }
            else {
                requestUrl = sharePointSite + "/_api/web/GetFolderById('" + this.folderId + "')/Files/AddUsingPath(DecodedUrl=@a1, overwrite='" + stringIsOverride + "', EnsureUniqueFileName='" + isVersionEnabled + "')?@a1='" + encodedFileName + "'&$expand=ListItemAllFields"
            }
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + decryptedToken,
                    "content-length": "0",
                    "Content-Type": "multipart/form-data"
                }),
            };

            this.executePost(requestUrl, "", httpOptions, selectedConnectorTab, selectedEntityConfiguration).then(fileResponse => {
                if (fileResponse.hasError == true) {
                    if ('comments' in fileResponse) {
                        if (fileResponse.comments.includes("2130575257") || fileResponse.comments.includes("tokenExpired")) {
                            // resolve(fileResponse);
                            resolve(fileResponse);
                        }
                        else {
                            reject(fileResponse);
                        }
                    }
                    else {
                        reject("error");
                        //resolve(true);
                    }
                }
                else {
                    // resolve(true);
                    resolve(fileResponse);

                }
            }
            ).catch((err: any) => {
                reject(err);
            });
        });
    }

    //the primary method that resursively calls to get the chunks and upload them to the library (to make the complete file)    
    uploadFileStart(result, id, serverRelativeURL, folderPath, fileName, itemID, fileUniqueId, chunks, index, byteOffset, chunkPercentage, resolve, reject, selectedConnectorTab, selectedEntityConfiguration, decryptedToken) {
        //we slice the file blob into the chunk we need to send in this request (byteOffset tells us the start position)    
        let self = this;
        let functionName: string = "uploadFileStart";
        let metaData: CRMEntity;
        const data = this.convertFileToBlobChunks(result, chunks[index]);
        //upload the chunk to the server using REST, using the unique upload guid as the identifier    
        this.uploadFileChunk(id, serverRelativeURL, fileName, fileUniqueId, chunks[index], data, byteOffset, selectedConnectorTab, selectedEntityConfiguration, decryptedToken).then(value => {
            const isFinished = index === chunks.length - 1;
            let response = {};
            let self = this;
            index += 1;
            const percentageComplete = isFinished ? 100 : Math.round((index * chunkPercentage));
            self.modalService.fileUploadingPercentage = `${percentageComplete}% Uploaded`;
            //console.log("Percentage Completed:" + percentageComplete);
            //More chunks to process before the file is finished, continue    
            if (index < chunks.length) {
                this.uploadFileStart(result, id, serverRelativeURL, folderPath, fileName, itemID, fileUniqueId, chunks, index, byteOffset, chunkPercentage, resolve, reject, selectedConnectorTab, selectedEntityConfiguration, decryptedToken);
            } else {
                if (selectedEntityConfiguration.isActivity && selectedEntityConfiguration.activityMetadataEnabled || selectedEntityConfiguration.linearMetadataEnabled) {
                    self.createSharePointMetadata$ = new Subject<any>();
                    self.createSharePointMetadataSub = self.createSharePointMetadata$.subscribe((resp) => {
                        response["FileName"] = fileName;
                        response["FilePath"] = serverRelativeURL;
                        response["status"] = "true";
                        response["FileUniqueId"] = fileUniqueId;
                        resolve(response);
                    })
                    self.createSharePointMetadata(selectedConnectorTab, selectedEntityConfiguration, itemID, folderPath);

                } else {
                    response["FileName"] = fileName;
                    response["FilePath"] = serverRelativeURL;
                    response["status"] = "true";
                    response["FileUniqueId"] = fileUniqueId;
                    resolve(response);
                }
                // response["FileName"] = fileName;
                // response["FilePath"] = serverRelativeURL;
                // response["status"] = "true";
                // resolve(response);
            }
        }).catch(error => {
            let response = {};
            response["FileName"] = fileName;
            response["FilePath"] = serverRelativeURL;
            response["status"] = "false";
            if (!self.a2dAppService.isValid(error.code)) {
                resolve(response);
            }
            else {
                self.spinnerService.hide();
                self.modalService.UploadStatusModalRef.hide();
                self.a2dAppService.logError('', JSON.stringify(error), selectedEntityConfiguration, '', fileName, serverRelativeURL);
                self.modalService.isOpen = true;
                self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
            }
        });
    }

    //this method sets up the REST request and then sends the chunk of file along with the unique indentifier (uploadId)    
    uploadFileChunk(id, serverRelativeURL, fileName, fileUniqueId, chunk, data, byteOffset, selectedConnectorTab, selectedEntityConfiguration, decryptedToken) {
        let functionName: string = "uploadFileChunk";
        return new Promise((resolve, reject) => {
            let httpOptions = {};
            let sharePointSite: string;
            let requestUrl: string;
            let fileServerRelativeURL: string;

            sharePointSite = selectedConnectorTab.absolute_url;

            let offset = chunk.offset === 0 ? '' : ',fileOffset=' + chunk.offset;

            //requestUrl = sharePointSite + "/_api/web/getfilebyserverrelativeurl('" + serverRelativeURL + "/" + fileName + "')/" + chunk.method + "(uploadId=guid'" + id + "'" + offset + ")";
            requestUrl = sharePointSite + "/_api/web/GetFileById('" + fileUniqueId + "')/" + chunk.method + "(uploadId=guid'" + id + "'" + offset + ")";
            httpOptions = {
                headers: new HttpHeaders({
                    "Accept": "application/json; odata=verbose",
                    "Content-Type": "application/octet-stream",
                    //'Content-Type': 'multipart/form-data',
                    "Authorization": "Bearer " + decryptedToken,
                }),
            };
            this.executePost(requestUrl, data, httpOptions, selectedConnectorTab, selectedEntityConfiguration).then(offset => {
                if (offset.hasError == true) {
                    if ('comments' in offset) {
                        if (offset.comments.includes("tokenExpired")) {
                            this.uploadFileChunk(id, serverRelativeURL, fileName, fileUniqueId, chunk, data, byteOffset, selectedConnectorTab, selectedEntityConfiguration, offset.newToken);
                        }
                        else {
                            reject(offset);
                        }
                    }
                }
                else {
                    resolve(offset);
                }
            }).catch(err => {
                reject(err);
            });
        });
    }

    //Helper method - depending on what chunk of data we are dealing with, we need to use the correct REST method...    
    getUploadMethod(offset, length, total) {
        if (offset + length + 1 > total) {
            return 'finishupload';
        } else if (offset === 0) {
            return 'startupload';
        } else if (offset < total) {
            return 'continueupload';
        }
        return null;
    }

    //this method slices the blob array buffer to the appropriate chunk and then calls off to get the BinaryString of that chunk    
    convertFileToBlobChunks(result, chunkInfo) {
        return result.slice(chunkInfo.offset, chunkInfo.offset + chunkInfo.length);
    }
    generateGUID() {
        function s4() {
            return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1);
        }
        return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
    }

    async executePost(url, data, requestHeaders, selectedConnectorTab, selectedEntityConfiguration) {
        let functionName: string = "executePost";
        const res = await this.http.post(url, data, requestHeaders).toPromise().catch(async (err: any) => {
            const error = err.error;
            let self = this;
            if ('error' in err.error) {
                //If file is alredy exist
                if (this.a2dAppService.isValid(err.error.error.code) && err.error.error.code.includes("2130575257")) {
                    //return this.parseRetSingle(error);
                    return err.error.error.code;
                }
                else {
                    return error;
                }
            }
            else if (this.a2dAppService.isValid(err.error.error_description) && (err.error.error_description = "Invalid JWT token. The token is expired.")) {
                const tokenResponse: any = await this.generateAccessTokensFromRefreshToken(selectedConnectorTab, selectedEntityConfiguration);
                this.acceessToken = tokenResponse.access_token;
                return {
                    tokenExpired: true,
                    newToken: tokenResponse.access_token
                };
            }
        });

        return this.parseRetSingle(res);
    }

    async executePatch(url, data, requestHeaders, selectedConnectorTab, selectedEntityConfiguration) {
        let functionName: string = "executePost";
        const res = await this.http.patch(url, data, requestHeaders).toPromise().catch(async (err: any) => {
            const error = err.error;
            let self = this;
            if ('error' in err.error) {
                //If file is alredy exist
                if (this.a2dAppService.isValid(err.error.error.code) && err.error.error.code.includes("2130575257")) {
                    //return this.parseRetSingle(error);
                    return err.error.error.code;
                }
                else {
                    return error;
                }
            }
            else if (this.a2dAppService.isValid(err.error.error_description) && (err.error.error_description = "Invalid JWT token. The token is expired.")) {
                const tokenResponse: any = await this.generateAccessTokensFromRefreshToken(selectedConnectorTab, selectedEntityConfiguration);
                this.acceessToken = tokenResponse.access_token;
                return {
                    tokenExpired: true,
                    newToken: tokenResponse.access_token
                };
            }
        });

        return this.parseRetSingle(res);
    }

    parseRetSingle(res) {
        if (res) {
            if (res.hasOwnProperty('d')) {
                return res.d;
            } else if (res.hasOwnProperty('error')) {
                const obj: any = res.error;
                obj.hasError = true;
                obj.comments = "error";
                return obj;
            }
            else if (res.hasOwnProperty('tokenExpired')) {
                const tokenObj: any = res;
                tokenObj.hasError = true;
                tokenObj.comments = "tokenExpired";
                return tokenObj;
            }
            else {
                return {
                    hasError: true,
                    comments: res
                };
            }
        } else {
            return {
                hasError: true,
                comments: 'Check the response in network trace'
            };
        }
    }
    //#endregion



    /**
* Validate the SharePoint Document Location in CRM and if it doesn't exist create one
* @param connector
* @param entityconfiguration
*/
    validateSPDocumentLocation(connector: Connector, entityconfiguration: EntityConfiguration, onChange?: any, folderPathArrayCol?: any): void {
        let functionName: string = "validateSPDocumentLocation";
        let self = this;
        let actionName: string = "ikl_SharePointCore";
        let folderPathArray: any = this.a2dAppService.isValid(folderPathArrayCol) ? folderPathArrayCol : [];
        // #Added 23/09/2019
        let updatedEntityConfiguration: EntityConfiguration = null;
        try {
            this.spinnerService.show();

            //this.CheckPasswordChanged(connector, entityconfiguration);


            this.retrieveDocumentLocations$ = new Subject<any>();
            this.validateFolder(connector, entityconfiguration);
            this.retrieveDocumentLocationsSub = this.retrieveDocumentLocations$.subscribe(
                (response) => {

                    // #Added 23/09/2019
                    if (this.a2dAppService.isValid(response) && this.a2dAppService.isValid(response["entityconfig"])) {
                        entityconfiguration = response["entityconfig"];
                        updatedEntityConfiguration = response["entityconfig"];
                    }

                    this.a2dAppService.documentLocationCountHome = this.a2dAppService.documentLocationCountHome + 1;
                    if (response['documentLocations'][0]["status"] == "false" || response['documentLocations'][0]["status"] == false) {
                        let entityName = this.a2dAppService.currentEntityName;
                        let recordId: any;
                        if (this.a2dAppService.selectedEntityRecords.length == 0) {
                            recordId = this.a2dAppService.currentEntityId;
                        }
                        else {
                            for (let z = 0; z < this.a2dAppService.selectedEntityRecords.length; z++) {
                                let id: any = this.a2dAppService.selectedEntityRecords[z].replace(/-/g, "");
                                if (entityconfiguration.folder_path.toLowerCase().indexOf(id.toLowerCase()) > -1) {
                                    recordId = id;
                                }
                            }
                        }

                        self.a2dAppService.validateFolderCreation$ = new Subject<any>();
                        self.validateFolderCreation(connector, entityconfiguration);
                        self.a2dAppService.validateFolderCreation$.subscribe(
                            (result) => {
                                if (this.a2dAppService.isValid(result)) {
                                    let validfolderStructure = this.a2dAppService.isValid(result["validfolderStructure"]) ? result["validfolderStructure"] : "";

                                    if (validfolderStructure.toLowerCase() == "true") {
                                        // #Added 23/09/2019
                                        /**
                                         * (1) Validate if Hierarchy Structure is enabled, if yes, then call SharePoint core 'createhierarchyfolders' which will create all the document locations
                                         * as per the hierarchy.
                                         * (2) Condition validates whether hierarchy is enabled, and if current record's (Interested in ABC {opportunity}) hierarchy record entity(A.Datum (account)) matches
                                         * with the SharePoint hierarchy enabled entity (i.e. either account or contact).
                                         */
                                        let object: any = {};

                                        // let connectorIndex  = self.a2dAppService.ConnectorList.findIndex((item)=>{return item.connector_id == response['connector'].connector_id});


                                        // response['connector'][] =  

                                        if (entityconfiguration.hierarchytype != "none" && this.a2dAppService.isValid(entityconfiguration.selectedHierarchyRecordType) && entityconfiguration.hierarchytype == entityconfiguration.selectedHierarchyRecordType) {
                                            //Create the Document location and also the folder in CRM (Hierarchy Structure)
                                            //Create the Document location and also the folder in CRM (Hierarchy Structure)
                                            if (this.a2dAppService.isSharePointSecuritySyncLicensePresent) {

                                                let grpPermissionId = this.a2dAppService.hierarchyGroupPermissionIdObj.filter((item) => { return item.entityConfigurationId == entityconfiguration.entity_configurationid });
                                                response['entityconfig']["hierarchyGroupId"] = this.a2dAppService.isValid(grpPermissionId) && this.a2dAppService.isValid(grpPermissionId[0]) && this.a2dAppService.isValid(grpPermissionId[0]["hierarchyGroupId"]) ? grpPermissionId[0]["hierarchyGroupId"] : "";
                                                response['entityconfig']["hirearchyPermissionId"] = this.a2dAppService.isValid(grpPermissionId) && this.a2dAppService.isValid(grpPermissionId[0]) && this.a2dAppService.isValid(grpPermissionId[0]["hirearchyPermissionId"]) ? grpPermissionId[0]["hirearchyPermissionId"] : "";
                                            }
                                            object = {
                                                "MethodName": "createhierarchyfolders",
                                                "ConnectorJSON": JSON.stringify(response['connector']),
                                                "CRMURL": self.a2dAppService._Xrm.getClientUrl(),
                                                "EntityName": entityName,
                                                "RecordId": recordId,
                                                "AdditionalDetailsJSON": JSON.stringify(response['documentLocations'][0]),
                                                "EntityConfigurationJSON": JSON.stringify(response['entityconfig'])
                                            }
                                        }
                                        else {
                                            //Create the Document location and also the folder in CRM (Linear Structure)
                                            object = {
                                                "MethodName": "createrecordfolder",
                                                "ConnectorJSON": JSON.stringify(response['connector']),
                                                "CRMURL": self.a2dAppService._Xrm.getClientUrl(),
                                                "EntityName": entityName,
                                                "RecordId": recordId,
                                                "AdditionalDetailsJSON": JSON.stringify(response['documentLocations'][0]),
                                                "EntityConfigurationJSON": JSON.stringify(response['entityconfig'])
                                            }
                                        }
                                        self.a2dAppService.callSharePointCoreAction(object, actionName, "DocumentLocation");
                                        self.actionOutputSub = self.a2dAppService.actionOutput$.subscribe(
                                            (response) => {
                                                // Create the Path based on which the Path would be created
                                                // #Added 23/09/2019
                                                /**
                                                 * Added condition, where if hierarchy structure is enabled for SharePoint, then no need of using parent document location since
                                                 * the entire hierarchy path comes from SharePoint Core as it has been handled that way.
                                                 */

                                                entityconfiguration = this.a2dAppService.isValid(response["entityconfig"]) ? response["entityconfig"] : (this.a2dAppService.isValid(updatedEntityConfiguration) ? updatedEntityConfiguration : entityconfiguration);
                                                connector = this.a2dAppService.isValid(response["connector"]) ? response["connector"] : connector;
                                                let path: string = "";
                                                if (entityconfiguration.hierarchytype != "none" && this.a2dAppService.isValid(entityconfiguration.selectedHierarchyRecordType) && entityconfiguration.hierarchytype == entityconfiguration.selectedHierarchyRecordType)
                                                    path = `/${response["DocumentLocation"]["relative_url"]}`;
                                                else
                                                    path = `/${response["DocumentLocation"]["parent_document_location_relative_url"]}/${response["DocumentLocation"]["relative_url"]}`;

                                                entityconfiguration.record_relative_url = path;

                                                this.a2dAppService._filteredDocumentLocationList[0] = this.a2dAppService.DocumentLocationList[0];

                                                this.a2dAppService._filteredDocumentLocationList[0]["path"] = path;
                                                for (let index = 0; index < self.a2dAppService.EntityConfigurationList.length; index++) {

                                                    // #Added 23/09/2019 - Added to set relative record url path
                                                    // Split based on forward slash, this is to pick up only the last two values, so that the things flow properly
                                                    let pathArr: string[] = path.split("/");
                                                    // Remove the starting slash from the path
                                                    path = path.startsWith("/") ? path.substr(1, path.length) : path;
                                                    // Get the record relative url
                                                    let recordRelativeURL: string = path;
                                                    // Added to handle hierarchy
                                                    if (entityconfiguration.hierarchytype != "none" && this.a2dAppService.isValid(entityconfiguration.selectedHierarchyRecordType) && entityconfiguration.hierarchytype == entityconfiguration.selectedHierarchyRecordType) {
                                                        // Validate if the length is greater than 2, and build the logic accordingly
                                                        if (pathArr.length > 2)
                                                            recordRelativeURL = pathArr.splice(pathArr.length - 2, pathArr.length - 1).join("/");
                                                        else
                                                            recordRelativeURL = pathArr.join("/");
                                                    }

                                                    if (self.a2dAppService.EntityConfigurationList[index].connector_type == self.a2dAppService.sharepoint) {
                                                        if (self.a2dAppService.EntityConfigurationList[index].connector_id == connector.connector_id) {
                                                            self.a2dAppService.EntityConfigurationList[index].folder_path = path;
                                                            self.a2dAppService.EntityConfigurationList[index].record_relative_url = recordRelativeURL;
                                                            self.a2dAppService.EntityConfigurationList[index].amILoaded = true;

                                                        }
                                                    }
                                                    // #Added 23/09/2019
                                                    // Update path in current entityConfiguration collection
                                                    if (entityconfiguration.connector_type == self.a2dAppService.sharepoint) {
                                                        if (entityconfiguration.connector_id == connector.connector_id) {
                                                            entityconfiguration.folder_path = path;
                                                            entityconfiguration.record_relative_url = recordRelativeURL;
                                                        }
                                                    }
                                                }
                                                if (this.a2dAppService.selectedEntityRecords.length == 0) {
                                                    this.a2dAppService.isSyncCompletedForSp = true;
                                                    //setTimeout(function () {
                                                    self.getSharePointData(connector, entityconfiguration, entityconfiguration.folder_path, self.gridService.selectedView);
                                                    // }, 5000);
                                                    //this.getSharePointData(connector, entityconfiguration, entityconfiguration.folder_path);
                                                    self.a2dAppService.setDefaultDropdownValue();
                                                }
                                                else {
                                                    if (this.a2dAppService.documentLocationCountHome == this.a2dAppService.selectedEntityRecords.length) {
                                                        this.a2dAppService.documentLocationCountHome = 0;
                                                        this.spinnerService.hide();
                                                    }
                                                }
                                            },
                                            (error) => {
                                                if (self.modalService.isOpen == false) {
                                                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                                }
                                                self.utilityService.throwError(error, functionName);
                                            }
                                        );
                                    }
                                }

                            },
                            (error) => {
                                self.a2dAppService.showEmptyDataMessage = true;
                                self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                self.utilityService.throwError(error, functionName);
                            }
                        )
                    }
                    else {
                        // if (this.a2dAppService.isSharePointSecuritySyncLicensePresent) {
                        //     let userId = this.a2dAppService._Xrm.getUserId().substring(1, this.a2dAppService._Xrm.getUserId().length - 1).toLowerCase();
                        //     this.a2dAppService.isSyncCompletedForSp = (this.a2dAppService.isValid(response.documentLocations) && (userId == response.documentLocations[0].ownerid)) ? true : response.documentLocations[0].isSync;
                        // } else {
                        //     this.a2dAppService.isSyncCompletedForSp = true;
                        // }
                        // Create the Path based on which the Path would be created
                        //let path: string = response["documentLocations"].find(function (x) { return x.is_active; }) ? response["documentLocations"].find(function (x) { return x.is_active; })["path"] : response["documentLocations"][0]["path"];

                        let path = this.a2dAppService._filteredDocumentLocationList.find(function (x) { return x.is_active; }) ? this.a2dAppService._filteredDocumentLocationList.find(function (x) { return x.is_active; })["path"] :
                            this.a2dAppService._filteredDocumentLocationList[0]["path"];

                        //let path = (this.a2dAppService._filteredDocumentLocationList.length !=0) ? this.a2dAppService._filteredDocumentLocationList.find(function (x) { return x.is_active; }) ? this.a2dAppService._filteredDocumentLocationList.find(function (x) { return x.is_active; })["path"] : 
                        //this.a2dAppService._filteredDocumentLocationList[0]["path"] : response["documentLocations"].find(function (x) { return x.is_active; }) ? response["documentLocations"].find(function (x) { return x.is_active; })["path"] : response["documentLocations"][0]["path"] ;

                        // Split based on forward slash, this is to pick up only the last two values, so that the things flow properly
                        let pathArr: string[] = path.split("/");
                        // Remove the starting slash from the path
                        path = path.startsWith("/") ? path.substr(1, path.length) : path;
                        // Get the record relative url
                        let recordRelativeURL: string = path;
                        // Validate if the length is greater than 2, and build the logic accordingly
                        if (pathArr.length > 2)
                            recordRelativeURL = pathArr.splice(pathArr.length - 2, pathArr.length - 1).join("/");
                        else
                            recordRelativeURL = pathArr.join("/");
                        this.a2dAppService._filteredDocumentLocationList[0]["path"] = path;
                        for (let index = 0; index < self.a2dAppService.EntityConfigurationList.length; index++) {
                            if (self.a2dAppService.EntityConfigurationList[index].connector_type == self.a2dAppService.sharepoint) {
                                if (self.a2dAppService.EntityConfigurationList[index].connector_id == connector.connector_id) {
                                    self.a2dAppService.EntityConfigurationList[index].amILoaded = true;
                                    self.a2dAppService.EntityConfigurationList[index].folder_path = path;
                                    self.a2dAppService.EntityConfigurationList[index].record_relative_url = recordRelativeURL;
                                }
                                entityconfiguration.folder_path = path;
                                entityconfiguration.record_relative_url = recordRelativeURL;
                            }
                        }

                        if (self.a2dAppService.isSharePointSecuritySyncLicensePresent && !(response.documentLocations[0].isSync)) {
                            let entityName = this.a2dAppService.currentEntityName;
                            let recordId: any;
                            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                                recordId = this.a2dAppService.currentEntityId;
                            }
                            else {
                                for (let z = 0; z < this.a2dAppService.selectedEntityRecords.length; z++) {
                                    let id: any = this.a2dAppService.selectedEntityRecords[z].replace(/-/g, "");
                                    if (entityconfiguration.folder_path.toLowerCase().indexOf(id.toLowerCase()) > -1) {
                                        recordId = id;
                                    }
                                }
                            }
                            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                                self.getSharePointData(connector, entityconfiguration, entityconfiguration.folder_path, self.gridService.selectedView);
                                self.a2dAppService.setDefaultDropdownValue();
                            }
                            else {
                                if (this.a2dAppService.documentLocationCountHome == this.a2dAppService.selectedEntityRecords.length) {
                                    this.a2dAppService.documentLocationCountHome = 0;
                                    this.spinnerService.hide();
                                }
                            }
                        } else {
                            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                                self.getSharePointData(connector, entityconfiguration, entityconfiguration.folder_path, self.gridService.selectedView);
                                self.a2dAppService.setDefaultDropdownValue();
                            }
                            else {
                                if (this.a2dAppService.documentLocationCountHome == this.a2dAppService.selectedEntityRecords.length) {
                                    this.a2dAppService.documentLocationCountHome = 0;
                                    this.spinnerService.hide();
                                }
                            }
                        }
                    }
                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }


    validateFolderCreation(connector: Connector, entityconfiguration: any) {
        let functionName = "validateFolderCreation";
        let self = this;
        let connectorJSON: string = "";
        try {
            //connectorJSON = JSON.stringify(connectorResponse);
            //Create the Document location and also the folder in CRM
            let data: any = {
                "MethodName": "validateFolderCreation",
                "EntityConfigurationJSON": JSON.stringify(entityconfiguration),
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": self.a2dAppService._Xrm.getClientUrl(),
                "EntityName": this.a2dAppService.currentEntityName,
            }
            //data: any, actionName: string, responseType: string, entityConfiguration?: any, connector?: any
            self.a2dAppService.callSharePointCoreAction(data, self.a2dAppService.spActionName, "ValidateFolderCreation", "", "", "");

        } catch (error) {
            self.a2dAppService.throwError(error, functionName);
        }
    }
    /**
      * Validate the SharePoint Document Location in CRM and if it doesn't exist create one
      * @param connector
      * @param entityconfiguration
      */
    validateSPDocumentLocationHome(connector: Connector, entityconfiguration: any): void {
        let functionName: string = "validateSPDocumentLocationHome";
        let self = this;
        let actionName: string = "ikl_SharePointCore";
        // #Added 23/09/2019
        let updatedEntityConfiguration: EntityConfiguration = null;
        try {
            this.spinnerService.show();
            this.retrieveDocumentLocations$ = new Subject<any>();
            this.validateFolder(connector, entityconfiguration);
            this.retrieveDocumentLocationsSub = this.retrieveDocumentLocations$.subscribe(
                (response) => {
                    if (this.a2dAppService.isValid(response) && this.a2dAppService.isValid(response["entityconfig"])) {
                        entityconfiguration = response["entityconfig"];
                        updatedEntityConfiguration = response["entityconfig"];
                    }

                    this.a2dAppService.documentLocationCountHome = this.a2dAppService.documentLocationCountHome + 1;
                    if (response['documentLocations'][0]["status"] == "false" || response['documentLocations'][0]["status"] == false) {
                        let entityName = this.a2dAppService.currentEntityName;
                        let recordId: any;
                        if (this.a2dAppService.selectedEntityRecords.length == 0) {
                            recordId = this.a2dAppService.currentEntityId;
                        }
                        else {
                            for (let z = 0; z < this.a2dAppService.selectedEntityRecords.length; z++) {
                                let id: any = this.a2dAppService.selectedEntityRecords[z].replace(/-/g, "");
                                if (entityconfiguration.folder_path.toLowerCase().indexOf(id.toLowerCase()) > -1) {
                                    recordId = id;
                                }
                            }
                        }
                        self.a2dAppService.validateFolderCreation$ = new Subject<any>();
                        self.validateFolderCreation(connector, entityconfiguration);
                        self.a2dAppService.validateFolderCreation$.subscribe(
                            (result) => {
                                if (this.a2dAppService.isValid(result)) {
                                    let validfolderStructure: string = this.a2dAppService.isValid(result["validfolderStructure"]) ? result["validfolderStructure"] : "";
                                    if (validfolderStructure.toLowerCase() == "true") {
                                        // #Added 23/09/2019
                                        // #Added 23/09/2019
                                        /**
                                         * (1) Validate if Hierarchy Structure is enabled, if yes, then call SharePoint core 'createhierarchyfolders' which will create all the document locations
                                         * as per the hierarchy.
                                         * (2) Condition validates whether hierarchy is enabled, and if current record's (Interested in ABC {opportunity}) hierarchy record entity(A.Datum (account)) matches
                                         * with the SharePoint hierarchy enabled entity (i.e. either account or contact).
                                        */

                                        let object: any = {};

                                        // Validate if Hierarchy type is set, and record's hierarchy matches with sharepoint hierarchy
                                        if (entityconfiguration.hierarchytype != "none" && this.a2dAppService.isValid(entityconfiguration.selectedHierarchyRecordType) && entityconfiguration.hierarchytype == entityconfiguration.selectedHierarchyRecordType) {
                                            //Create the Document location and also the folder in CRM (Hierarchy Structure)
                                            if (this.a2dAppService.isSharePointSecuritySyncLicensePresent) {

                                                let grpPermissionId = this.a2dAppService.hierarchyGroupPermissionIdObj.filter((item) => { return item.entityConfigurationId == entityconfiguration.entity_configurationid });
                                                response['entityconfig']["hierarchyGroupId"] = this.a2dAppService.isValid(grpPermissionId) && this.a2dAppService.isValid(grpPermissionId[0]) && this.a2dAppService.isValid(grpPermissionId[0]["hierarchyGroupId"]) ? grpPermissionId[0]["hierarchyGroupId"] : "";
                                                response['entityconfig']["hirearchyPermissionId"] = this.a2dAppService.isValid(grpPermissionId) && this.a2dAppService.isValid(grpPermissionId[0]) && this.a2dAppService.isValid(grpPermissionId[0]["hirearchyPermissionId"]) ? grpPermissionId[0]["hirearchyPermissionId"] : "";
                                            }

                                            object = {
                                                "MethodName": "createhierarchyfolders",
                                                "ConnectorJSON": JSON.stringify(response['connector']),
                                                "CRMURL": self.a2dAppService._Xrm.getClientUrl(),
                                                "EntityName": entityName,
                                                "RecordId": recordId,
                                                "AdditionalDetailsJSON": JSON.stringify(response['documentLocations'][0]),
                                                "EntityConfigurationJSON": JSON.stringify(response['entityconfig'])
                                            }
                                        }
                                        else {
                                            //Create the Document location and also the folder in CRM (Linear Structure)
                                            object = {
                                                "MethodName": "createrecordfolder",
                                                "ConnectorJSON": JSON.stringify(response['connector']),
                                                "CRMURL": self.a2dAppService._Xrm.getClientUrl(),
                                                "EntityName": entityName,
                                                "RecordId": recordId,
                                                "AdditionalDetailsJSON": JSON.stringify(response['documentLocations'][0]),
                                                "EntityConfigurationJSON": JSON.stringify(response['entityconfig'])
                                            }
                                        }
                                        self.a2dAppService.callSharePointCoreAction(object, actionName, "DocumentLocation", response['entityconfig']);
                                        self.actionOutputSub = self.a2dAppService.actionOutput$.subscribe(
                                            (response) => {
                                                // Create the Path based on which the Path would be created
                                                //let path: string = `/${response["DocumentLocation"]["parent_document_location_relative_url"]}/${response["DocumentLocation"]["relative_url"]}`;

                                                // Create the Path based on which the Path would be created

                                                // #Added 23/09/2019
                                                /**
                                                 * Added condition, where if hierarchy structure is enabled for SharePoint, then no need of using parent document location since
                                                 * the entire hierarchy path comes from SharePoint Core as it has been handled that way.
                                                 */

                                                entityconfiguration = this.a2dAppService.isValid(response["entityconfig"]) ? response["entityconfig"] : (this.a2dAppService.isValid(updatedEntityConfiguration) ? updatedEntityConfiguration : entityconfiguration);
                                                connector = this.a2dAppService.isValid(response["connector"]) ? response["connector"] : connector;
                                                let path: string = "";
                                                let entityConfig: any = response["entityconfig"];
                                                if (entityConfig.hierarchytype != "none" && this.a2dAppService.isValid(entityConfig.selectedHierarchyRecordType) && entityConfig.hierarchytype == entityConfig.selectedHierarchyRecordType)
                                                    path = `/${response["DocumentLocation"]["relative_url"]}`;
                                                else
                                                    path = `/${response["DocumentLocation"]["parent_document_location_relative_url"]}/${response["DocumentLocation"]["relative_url"]}`;

                                                entityconfiguration.record_relative_url = path;

                                                //entityconfiguration.folder_path = path;

                                                // #Added 23/09/2019
                                                this.a2dAppService._filteredDocumentLocationList[0] = this.a2dAppService.DocumentLocationList[0];

                                                this.a2dAppService._filteredDocumentLocationList[0]["path"] = path;
                                                response["entityconfig"].amILoaded = true;
                                                response["entityconfig"].folder_path = path;
                                                let pathArr: string[] = path.split("/");
                                                // Remove the starting slash from the path
                                                let recordRelativeURL: string = path.startsWith("/") ? path.substr(1, path.length) : path;

                                                // Validate if the length is greater than 2, and build the logic accordingly
                                                if (pathArr.length > 2)
                                                    recordRelativeURL = pathArr.splice(pathArr.length - 2, pathArr.length - 1).join("/");
                                                else
                                                    recordRelativeURL = pathArr.join("/");
                                                response["entityconfig"].record_relative_url = recordRelativeURL;

                                                // #Added 23/09/2019
                                                for (let index = 0; index < self.a2dAppService.EntityConfigurationList.length; index++) {
                                                    if (self.a2dAppService.EntityConfigurationList[index].connector_type == self.a2dAppService.sharepoint) {
                                                        if (self.a2dAppService.EntityConfigurationList[index].connector_id == connector.connector_id && self.a2dAppService.EntityConfigurationList[index].currentRecordId == entityconfiguration.currentRecordId) {
                                                            self.a2dAppService.EntityConfigurationList[index].folder_path = path;
                                                            self.a2dAppService.EntityConfigurationList[index].record_relative_url = recordRelativeURL;
                                                            self.a2dAppService.EntityConfigurationList[index].amILoaded = true;

                                                        }
                                                    }
                                                    // #Added 23/09/2019
                                                    // Update path in current entityConfiguration collection
                                                    if (entityconfiguration.connector_type == self.a2dAppService.sharepoint) {
                                                        if (entityconfiguration.connector_id == connector.connector_id) {
                                                            entityconfiguration.folder_path = path;
                                                        }
                                                    }
                                                }

                                                if (this.a2dAppService.selectedEntityRecords.length == 0) {

                                                    self.a2dAppService.isSyncCompletedForSp = true;

                                                    self.getSharePointData(connector, entityconfiguration, entityconfiguration.folder_path, self.gridService.selectedView);

                                                    self.a2dAppService.setDefaultDropdownValue();
                                                    ///this.spinnerService.hide();
                                                }
                                                else {
                                                    if (this.a2dAppService.documentLocationCountHome == this.a2dAppService.selectedEntityRecords.length) {
                                                        this.a2dAppService.documentLocationCountHome = 0;
                                                        this.spinnerService.hide();
                                                    }
                                                }
                                            },
                                            (error) => {
                                                if (self.modalService.isOpen == false) {
                                                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                                }
                                                self.utilityService.throwError(error, functionName);
                                            }
                                        );
                                    }
                                }
                            },
                            (error) => {
                                self.a2dAppService.showEmptyDataMessage = true;
                                self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                self.utilityService.throwError(error, functionName);
                            });
                    }
                    else {


                        // Create the Path based on which the Path would be created
                        // let path: string = response["documentLocations"].find(function (x) { return x.is_active; }) ? response["documentLocations"].find(function (x) { return x.is_active; })["path"] : response["documentLocations"][0]["path"];

                        let path = this.a2dAppService._filteredDocumentLocationList.find(function (x) { return x.is_active; }) ? this.a2dAppService._filteredDocumentLocationList.find(function (x) { return x.is_active; })["path"] :
                            this.a2dAppService._filteredDocumentLocationList[0]["path"];

                        // Split based on forward slash, this is to pick up only the last two values, so that the things flow properly
                        let pathArr: string[] = path.split("/");
                        // Remove the starting slash from the path
                        path = path.startsWith("/") ? path.substr(1, path.length) : path;
                        // Get the record relative url
                        let recordRelativeURL: string = path;
                        // Validate if the length is greater than 2, and build the logic accordingly
                        if (pathArr.length > 2)
                            recordRelativeURL = pathArr.splice(pathArr.length - 2, pathArr.length - 1).join("/");
                        else
                            recordRelativeURL = pathArr.join("/");
                        this.a2dAppService._filteredDocumentLocationList[0]["path"] = path;
                        response["entityconfig"].amILoaded = true;
                        response["entityconfig"].folder_path = path;
                        response["entityconfig"].record_relative_url = recordRelativeURL;

                        if (self.a2dAppService.isSharePointSecuritySyncLicensePresent && !(response.documentLocations[0].isSync)) {
                            let entityName = this.a2dAppService.currentEntityName;
                            let recordId: any;
                            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                                recordId = this.a2dAppService.currentEntityId;
                            }
                            else {
                                for (let z = 0; z < this.a2dAppService.selectedEntityRecords.length; z++) {
                                    let id: any = this.a2dAppService.selectedEntityRecords[z].replace(/-/g, "");
                                    if (entityconfiguration.folder_path.toLowerCase().indexOf(id.toLowerCase()) > -1) {
                                        recordId = id;
                                    }
                                }
                            }
                            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                                self.getSharePointData(connector, entityconfiguration, entityconfiguration.folder_path, self.gridService.selectedView);
                                self.a2dAppService.setDefaultDropdownValue();
                            }
                            else {
                                if (this.a2dAppService.documentLocationCountHome == this.a2dAppService.selectedEntityRecords.length) {
                                    this.a2dAppService.documentLocationCountHome = 0;
                                    this.spinnerService.hide();
                                }
                            }
                        } else {
                            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                                self.getSharePointData(connector, entityconfiguration, entityconfiguration.folder_path, self.gridService.selectedView);
                                self.a2dAppService.setDefaultDropdownValue();
                            }
                            else {
                                if (this.a2dAppService.documentLocationCountHome == this.a2dAppService.selectedEntityRecords.length) {
                                    this.a2dAppService.documentLocationCountHome = 0;
                                    this.spinnerService.hide();
                                }
                            }
                        }


                    }
                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Validate folder or else create one if it doesn't exist
     * @param connector
     * @param entityConfiguration
     * return void
     */
    validateFolder(connector: Connector, entityConfiguration: EntityConfiguration): void {
        let functionName: string = "validateFolder";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        // #Added 23/09/2019
        let proceedResponse: boolean = true;
        try {
            entityName = this.a2dAppService.currentEntityName;

            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                recordId = this.a2dAppService.currentEntityId;
            }
            else {
                for (let z = 0; z < this.a2dAppService.selectedEntityRecords.length; z++) {
                    let id: any = this.a2dAppService.selectedEntityRecords[z].replace(/-/g, "");
                    if (entityConfiguration.folder_path.indexOf(id) > -1) {
                        recordId = id;
                    }
                }
            }

            this.a2dAppService.retrieveDocumentLocations$ = new Subject<any>();
            this.a2dAppService.retrieveDocumentLocation(recordId, entityName, connector.sharepoint_site_id, connector, entityConfiguration, null, null, 0);
            this.retrieveDocumentLocationsSub = this.a2dAppService.retrieveDocumentLocations$.subscribe(
                (response) => {
                    if (!this.a2dAppService.isValid(response['documentLocations'])) {
                        self.a2dAppService.retrieveEntityDefinitions('', 'Root Document Location is Not present for ' + entityConfiguration.rootEntityDisplayName + ' entity.' + ' - { From validateFolder }', entityConfiguration, '', '');
                        self.a2dAppService.entityDefinitionSub = self.a2dAppService.entityDefinitions$.subscribe(
                            (response) => {
                                self.a2dAppService.objectTypeCodeArray[self.a2dAppService.currentEntityName.toLowerCase()] = response[0].ObjectTypeCode;
                                self.createErrorLog(response["errorResponse"], entityConfiguration, response[0].ObjectTypeCode);
                                self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                            }
                        );
                    }
                    else {
                        let responseNew = response['documentLocations'] as DocumentLocation[];

                        // If the parent_document_location_id is empty, which means that the current record doesn't have a document location record
                        // The conclusion is based off of the logic that has been implemented in retrieving the document location
                        if (this.a2dAppService.isValid(responseNew) && this.a2dAppService.isValid(responseNew[0]) && responseNew[0]["parent_document_location_id"]) {
                            responseNew[0]["status"] = true;
                        }
                        // If status is false then show the error message, this will occur when the Document location record is not there for,
                        // even the entity, i.e., 'account' doesn't have a document location record
                        else if (this.a2dAppService.isValid(response["status"]) && response["status"] == false) {
                            self.modalService.openErrorDialog(self.message_wronglyConfiguredEC, (onOKClick) => { })
                        }
                        // #Added 23/09/2019
                        /**
                         * (1) Condition has been added specific for hierarchy strucuture, where if hierarchy record's entity document
                         * location is not present (i.e. hierarchy record is 'A.Datum', its entity document location is 'Account',
                         * and current record is 'Interested in ABC (opportunity'), it will create hierarchy record's entity document location in CRM (i.e. 'Account')
                         * (2) In this case, we'll set the proceedResponse flag to false, so that the subscriber isn't called to perform further operations, and will call validateFolder
                         * method again to follow hierarchical process.
                         */
                        // Check  conditions to validate if hierarchy record's entity document location needs to be created or not.
                        else if (this.a2dAppService.isValid(response["status"]) && response["status"] == "false" &&
                            this.a2dAppService.isValid(response["IsHierarchy"]) && response["IsHierarchy"] == true &&
                            this.a2dAppService.isValid(response["HierarchyType"]) && response["HierarchyType"] != "none" &&
                            this.a2dAppService.isValid(response["HierarchyRecordId"])) {
                            let connectorData = response["ConnectorData"];
                            let entityConfigurationData = response["EntityConfigurationData"];
                            proceedResponse = false;
                            // Call SharePoint core to create hierarchy record parent entity (account or contact)
                            this.validateFolderInSharePoint(connectorData, entityConfigurationData, response["HierarchyType"], response["HierarchyRecordId"])
                            this.validateFolderInSharePointSub = this.a2dAppService.validateFolderInSharePoint$.subscribe(
                                (response) => {
                                    // If response is true, then recall validateFolder to follow hierarchy process.
                                    if (this.a2dAppService.isValid(response) && response["status"] == true) {
                                        let connectorObj = response["connector"];
                                        let entityConfigObj = response["entityConfiguration"];
                                        self.validateFolder(connectorObj, entityConfigObj);
                                    }
                                },
                                (error) => {
                                    responseNew[0]["status"] = false;
                                });
                        }
                        else {
                            responseNew[0]["status"] = false;
                        }

                        // #Added 23/09/2019
                        /**
                         * (1) Don't proceed with response if false.
                         * (2) proceedResponse will be false, when SharePoint hierarchy structure is enabled, and current record's hierarchy record's parent entity's
                         * document location is not created.
                         * (3) E.g.: If Hierarchy is enabled based on 'Account' and current record is Interested in ABC (Opportunity), and it's hierarchy is set as A.Datum (account)
                         * if, A.Datum(account) document location is not present in CRM, then we'll not proceed ahead.
                         */
                        if (proceedResponse) {
                            response['documentLocations'] = responseNew;
                            this.retrieveDocumentLocations$.next(response);
                        }
                    }

                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }


    validateFolderInSharePoint(connector: Connector, entityConfiguration: EntityConfiguration, entityName: string, hierarchyRecordId: string): void {
        let functionName: string = "validateFolderInSharePoint";

        let actionName: string = "ikl_SharePointCore";
        let connectorJSON = "";
        let additionalDetailsJSON = "";
        let additionalDetails: {} = {};
        let data: {} = {};
        try {
            //Get the JSON out of the conenctObject
            connectorJSON = JSON.stringify(connector);

            //Add the entity name
            additionalDetails["entity_validate_document_library"] = entityName;
            additionalDetailsJSON = JSON.stringify(additionalDetails);

            //Create the Data
            data = {
                "MethodName": "validatefolder",
                "ConnectorJSON": connectorJSON,
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": hierarchyRecordId,
                "AdditionalDetailsJSON": additionalDetailsJSON
            };

            this.a2dAppService.callSharePointCoreAction(data, actionName, "ValidateFolderInSharePoint", entityConfiguration, connector);

        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
      * This function performs get data from sharepoint
      * @param connector
      * @param entityConfiguration
      * @param path
      */
    getSharePointData(connector: Connector, entityConfiguration: EntityConfiguration, path: string, viewId?: string, rowData?: any, onSaveFunction?: string) {
        let functionName: string = "getSharePointData";
        let self = this;
        try {
            this.a2dAppService.showEmptyDataMessage = false;
            if (onSaveFunction != "onScroll") {
                this.spinnerService.show();
            }

            this.getFiles(connector, entityConfiguration, path, viewId, onSaveFunction);
            this.fillGridSub = this.getFiles$.subscribe(
                (result) => {
                    this.a2dAppService.showEmptyDataMessage = result.length > 0 ? false : true;

                    if (self.utilityService.gridData.length > 0 && onSaveFunction == "onScroll") {
                        //self.utilityService.gridData = [...self.utilityService.gridData, result]
                        self.utilityService.gridData = self.utilityService.gridData.concat(result);
                    } else {
                        self.utilityService.gridData = result;
                    }
                    //self.getNextRows();
                    self.breadcrumbService.getDataForBreadcrumb(connector, path, '', entityConfiguration);
                    self.gridService.sharePointService = self;
                    self.fillGridSub.unsubscribe();
                    self.spinnerService.hide();
                },
                (error) => {
                    self.a2dAppService.showEmptyDataMessage = true;
                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    self.utilityService.throwError(error, functionName);
                }
            )
            //}, 5000);
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Function perform show and hide of dropdown for document locations
     * @param connector
     * return boolean
     */
    hideShowDropDown(connector: Connector): boolean {
        let functionName: string = "hideShowDropDown";
        let showDropdown: boolean = false;
        try {
            if (this.a2dAppService.isValid(connector)) {
                if (this.a2dAppService.selectedEntityRecords.length > 0) {
                    showDropdown = false;
                }
                else {
                    if (connector.connector_type_value == this.a2dAppService.sharepoint) {
                        if (this.a2dAppService.currentEntityName.toLowerCase() != 'email' || this.a2dAppService.allowActivityFolderCreation) {
                            showDropdown = true;
                            this.a2dAppService.setDefaultDropdownValue();
                        }
                        else {
                            showDropdown = false;
                        }
                    }
                    else {
                        showDropdown = false;
                    }
                }
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return showDropdown;
    }

    /**
     * Get Files from SharePoint
     * @param connector
     * @param entityConfiguration
     * @param path
     */
    getFiles(connector: Connector, entityConfiguration: EntityConfiguration, path: string, viewId?: string, callingFunction?: string) {
        let functionName: string = "getFiles";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let additionalDetails: DocumentLocation = {};
        try {
            //this.spinnerService.show();
            this.getFilesAngularSub = new Subscription();
            this.getFiles$ = new Subject<any>();
            this.getMainRootLibraryViews$ = new Subject<any>();
            this.createViewXml$ = new Subject<any>();

            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;

            if (this.a2dAppService._filteredDocumentLocationList.length > 0) {
                additionalDetails = this.a2dAppService._filteredDocumentLocationList[0]; //Get the first out of it
            }
            // Get the SharePoint Sub Site if any
            let subSite: string = this.utilityService.getSharePointSubSite(connector.absolute_url);
            //get the cleared path
            path = this.utilityService.clearSubSiteFromPath(subSite, path);
            additionalDetails["current_path"] = path;
            //Create the Document location and also the folder in CRM
            this.createViewXml(connector, entityConfiguration, path, viewId);
            this.createViewXmlSub = this.createViewXml$.subscribe((viewXml) => {
                self.getFilesAngularSub = new Subscription();
                self.getFilesAngular$ = new Subject<any>();
                self.getFilesAngular(connector, entityConfiguration, path, viewId, viewXml, callingFunction);
                self.getFilesAngularSub = this.getFilesAngular$.subscribe((getFilesResponse: any) => {

                    self.a2dAppService.columnsArray = getFilesResponse.ListSchema.Field;
                    this.folderId = getFilesResponse.ListData.CurrentFolderUniqueId;
                    // if (self.a2dAppService.isValid(viewId)) {
                    // if (self.gridService.selectedView != viewId)
                    self.gridService.cols = [];
                    // }
                    if (self.gridService.cols.length == 0) {
                        self.a2dAppService.columnsArray.forEach(element => {
                            if (element.RealFieldName != 'DocIcon' /*&& element.FieldType != "User"*/ && element.RealFieldName != 'FileSizeDisplay') {
                                self.gridService.cols.push({
                                    field: element.RealFieldName, header: element.DisplayName, width: "2rem",
                                    fieldType: element.FieldType.toLowerCase() != 'datetime' ? element.FieldType.toLowerCase() : element.Format.toLowerCase(),
                                    choices: this.a2dAppService.isValid(element.Choices) ? element.Choices : null,
                                    isReadOnlyField: this.a2dAppService.isValid(element.ReadOnly) ? Boolean(element.ReadOnly.toLowerCase()) : false,
                                    listId: element.FieldType.toLowerCase() == 'lookup' && element.DispFormUrl ? new URLSearchParams(new URL(element.DispFormUrl).search).get('ListId') : ""
                                });
                            }
                        });
                        self.gridService.cols.push({
                            field: "File_x0020_Size", header: self.a2dAppService.labelsMultiLanguage['thsize'], width: "2rem",
                            fieldType: "number",
                            choices: null,
                            isReadOnlyField: true
                        }
                        );
                    }
                    self.nextRowDataHref = self.a2dAppService.isValid(getFilesResponse.ListData.NextHref) ? getFilesResponse.ListData.NextHref.split('?')[1] : "";
                    let files: any = self.createCollectionSharePointData(getFilesResponse.ListData.Row, getFilesResponse.ListSchema.Field, connector, entityConfiguration);

                    //self.spinnerService.hide();
                    self.getFiles$.next(files);
                });
            })

        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
 * This function performs convert numbers to comma separated
 * @param size 
 */
    Convert(size: any): any {
        let functionName: string = "Convert";
        try {
            if (this.a2dAppService.isValid(size)) {
                size = parseFloat(size);
                return size.toFixed(2);
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Method to get the files and the columns and views of SharePoint
     * @param connector 
     * @param entityConfiguration 
     * @param path 
     * @param viewId 
     * @param ViewXml 
     */
    getFilesAngular(connector: Connector, entityConfiguration: EntityConfiguration, path: string, viewId?: string, ViewXml?: string, callingFunction?: string, rowData?: any, fileName?: string) {
        let functionName: string = "getFilesAngular";
        try {
            //return new Promise((resolve, reject) => {
            // Construct the endpoint - The GetList method is available for SharePoint Online only. 
            let self = this;
            let httpOptions = {};
            let stringIsOverride: string;
            let sharePointSite: string;
            let requestUrl: string;
            let data: string;
            let rootFolderName: string;
            let serverRelativeURL: string;
            let folderRelativePath: string;
            let sitePath: any;
            let tempSitePath: any;
            if (callingFunction != 'onScroll') {
                this.a2dAppService.selectedData = []; // selected Fill become null 
            }            //No need this line because we're passing it in url not in (parenthesis)
            //folderRelativePath = path.replace(/'/g, "''");
            folderRelativePath = path;
            viewId = "";
            //@ts-ignore
            this.acceessToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;

            rootFolderName = entityConfiguration.folder_path.replace(/^\/+|\/+$/g, '').split('/')[0];
            sitePath = connector.absolute_url.split("/");
            tempSitePath = connector.absolute_url.split("/");
            //Checking if path have "/"" character at first position if contains then remove because it is not supports inrequest url (GetFolderByServerRelativeUrl).
            if (sitePath.length > 3 && !folderRelativePath.includes(tempSitePath.splice(3).join('/'))) {
                sitePath = '/' + sitePath.splice(3).join('/');
                folderRelativePath = sitePath + "/" + folderRelativePath;
                serverRelativeURL = encodeURIComponent(folderRelativePath.replace(/\/+/g, '/'));
                rootFolderName = encodeURIComponent(sitePath + '/' + rootFolderName);
            }
            else {
                if (folderRelativePath.includes(tempSitePath.splice(3).join('/'))) {
                    sitePath = '/' + sitePath.splice(3).join('/');
                    rootFolderName = encodeURIComponent(sitePath + '/' + rootFolderName);
                }
                serverRelativeURL = encodeURIComponent("/" + folderRelativePath.replace(/^\/+|\/+$/g, "").replace(/\/+/g, "/"));
            }
            sharePointSite = connector.absolute_url;

            if (this.a2dAppService.isValid(this.nextRowDataHref) && callingFunction == 'onScroll') {
                requestUrl = sharePointSite + "/_api/web/GetListUsingPath(decodedurl='" + rootFolderName + "')/RenderListDataAsStream?&View=" + viewId + "&TryNewExperienceSingle=TRUE&" + this.nextRowDataHref;

            }
            else if (callingFunction == "updateSavedRowData" && this.a2dAppService.isValid(rowData)) {
                requestUrl = sharePointSite + "/_api/web/GetListUsingPath(decodedurl='" + rootFolderName + "')/RenderListDataAsStream?&RootFolder=" + serverRelativeURL + "&TryNewExperienceSingle=TRUE?&View=" + viewId + "&FilterField1=UniqueId&FilterValue1=" + rowData["UniqueId"] + "";
            }
            else if (callingFunction == "SearchFiles" && this.a2dAppService.isValid(fileName)) {
                requestUrl = sharePointSite + "/_api/web/GetListUsingPath(decodedurl='" + rootFolderName + "')/RenderListDataAsStream?&RootFolder=" + serverRelativeURL + "&?&View=" + viewId + "TryNewExperienceSingle=TRUE&InplaceSearchQuery=" + fileName;
            }
            else {
                requestUrl = sharePointSite + "/_api/web/GetListUsingPath(decodedurl='" + rootFolderName + "')/RenderListDataAsStream?&RootFolder=" + serverRelativeURL + "&View=" + viewId + "&TryNewExperienceSingle=TRUE";
            }

            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + this.acceessToken,
                    "Content-Type": "application/json;odata=verbose"
                }),
            };
            // if (callingFunction != "updateSavedRowData" && callingFunction != "SearchFiles") {
            if (!this.a2dAppService.isValid(ViewXml)) {
                data = JSON.stringify({ "parameters": { "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": 7837447, "AllowMultipleValueFilterForTaxonomyFields": true, "AddRequiredFields": true, "RequireFolderColoringFields": true } }); //"ViewXml":"<RowLimit Paged=\"TRUE\">1000</RowLimit>", //,"ViewXml":"<View><RowLimit Paged=\"TRUE\">30</RowLimit><QueryOptions><Paging ListItemCollectionPositionNext=\"\"/></QueryOptions></View>"
            }
            else {
                data = JSON.stringify({ "parameters": { "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": 7837447, "AllowMultipleValueFilterForTaxonomyFields": true, "AddRequiredFields": true, "RequireFolderColoringFields": true, "ViewXml": ViewXml } }); //"ViewXml":"<RowLimit Paged=\"TRUE\">1000</RowLimit>", //,"ViewXml":"<View><RowLimit Paged=\"TRUE\">30</RowLimit><QueryOptions><Paging ListItemCollectionPositionNext=\"\"/></QueryOptions></View>"
            }
            // }
            this.executePost(requestUrl, data, httpOptions, connector, entityConfiguration).then(response => {
                if (response.hasError == true && typeof response.comments == 'string') {
                    if ('comments' in response) {
                        if (response.comments.includes("2130575257") || response.comments.includes("tokenExpired")) {
                            // resolve(fileResponse);
                            // console.log("Response of " + functionName + ": ", response);
                            self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then(
                                (response: any) => {
                                    this.acceessToken = self.a2dAppService.isValid(response) ? response.access_token : null;
                                    self.getFilesAngular(connector, entityConfiguration, path, viewId);
                                },
                                (error: any) => {
                                    self.getFilesAngular$.next(error);
                                }
                            );
                        }
                        else {
                            if (self.modalService.isOpen == false) {
                                if (response.message) {
                                    self.a2dAppService.retrieveEntityDefinitions('', response.message.value + ' - { From GetFiles }', entityConfiguration, '', '');
                                    self.a2dAppService.entityDefinitionSub = self.a2dAppService.entityDefinitions$.subscribe(
                                        (response) => {
                                            self.a2dAppService.objectTypeCodeArray[self.a2dAppService.currentEntityName.toLowerCase()] = response[0].ObjectTypeCode;
                                            self.createErrorLog(response["errorResponse"], entityConfiguration, response[0].ObjectTypeCode, path);
                                        }
                                    );
                                }
                                self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                            }
                        }
                    }
                    else {
                        if (self.modalService.isOpen == false) {
                            if (response.message) {
                                self.a2dAppService.retrieveEntityDefinitions('', response.message.value + ' - { From GetFiles }', entityConfiguration, '', '');
                                self.a2dAppService.entityDefinitionSub = self.a2dAppService.entityDefinitions$.subscribe(
                                    (response) => {
                                        self.a2dAppService.objectTypeCodeArray[self.a2dAppService.currentEntityName.toLowerCase()] = response[0].ObjectTypeCode;
                                        self.createErrorLog(response["errorResponse"], entityConfiguration, response[0].ObjectTypeCode, path);
                                    }
                                );
                            }
                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                        self.getFilesAngular$.next(response);
                    }
                }
                else {
                    // console.log("Response of " + functionName + ": ", response);
                    self.getFilesAngular$.next(response.comments);

                }
            }
            ).catch((err: any) => {
                self.getFilesAngular$.next("false");
            });
            //});
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }



    getMainRootLibraryViews(connector: Connector, entityConfiguration: EntityConfiguration, colOrView?: string) {
        let functionName: string = "getMainRootLibraryViews";
        let self = this;
        let requestUrl: string;
        let httpOptions: any;
        let sitePath: any;
        let tempSitePath: any;
        try {

            let rootFolderName = entityConfiguration.folder_path.replace(/^\/+|\/+$/g, '').split('/')[0];
            sitePath = connector.absolute_url.split("/");
            tempSitePath = connector.absolute_url.split("/");
            //Checking if path have "/"" character at first position if contains then remove because it is not supports inrequest url (GetFolderByServerRelativeUrl).
            if (sitePath.length > 3) {
                sitePath = '/' + sitePath.splice(3).join('/');
                rootFolderName = encodeURIComponent(sitePath + '/' + rootFolderName);
            }
            else {
                rootFolderName = encodeURIComponent(rootFolderName);

            }
            if (this.a2dAppService.isValid(colOrView) && colOrView == "views") {
                requestUrl = connector.absolute_url + "/_api/web/lists/getbytitle('" + entityConfiguration.rootEntityDisplayName + "')/views";
            }
            else {
                requestUrl = connector.absolute_url + "/_api/web/getList('" + rootFolderName + "')/fields";
            }
            //@ts-ignore
            this.acceessToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + this.acceessToken,
                    "Content-Type": "application/json;odata=verbose"
                }),
            };
            this.http.get(requestUrl, httpOptions,).subscribe((resp) => {
                self.getMainRootLibraryViews$.next(resp);
            }, (err) => {
                if (err && err.error && err.error.error_description && err.error.error_description.includes("Invalid JWT token. The token is expired.")) {
                    self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then((tokenResponse: any) => {
                        self.acceessToken = tokenResponse ? tokenResponse.access_token : null;

                        if (tokenResponse && self.a2dAppService.isValid(tokenResponse.access_token)) {
                            self.getMainRootLibraryViews(connector, entityConfiguration, colOrView);
                        }
                    })

                }
                if (self.modalService.isOpen == false) {
                    if (err.error.error.message) {
                        self.a2dAppService.retrieveEntityDefinitions('', err.error.error.message.value + ' - { From getMainRootLibraryViews }', entityConfiguration, '', '');
                        self.a2dAppService.entityDefinitionSub = self.a2dAppService.entityDefinitions$.subscribe(
                            (response) => {
                                self.a2dAppService.objectTypeCodeArray[self.a2dAppService.currentEntityName.toLowerCase()] = response[0].ObjectTypeCode;
                                self.createErrorLog(response["errorResponse"], entityConfiguration, response[0].ObjectTypeCode, requestUrl);
                            }
                        );
                    }
                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                }
            }
            )


        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    createViewXml(connector: Connector, entityConfiguration: EntityConfiguration, path: string, viewId: string) {
        let functionName: string = "createViewXml";
        let field: any;
        let selectedView: any;
        let viewXml: string;
        let self = this;
        try {
            this.getFilesAngular$ = new Subject<any>();
            this.getMainRootLibraryViews$ = new Subject<any>();
            if (!this.a2dAppService.isValid(viewId) && this.a2dAppService.Views.length == 0 && !this.a2dAppService.isValid(this.gridService.selectedViewThumbnail)) {
                this.getMainRootLibraryViews(connector, entityConfiguration, "views")
                this.getMainRootLibraryViewsSub = this.getMainRootLibraryViews$.subscribe((getViewsResponse) => {
                    self.a2dAppService.Views = getViewsResponse.d.results.filter(view => view.Title != 'Merge Documents' && view.Title != "assetLibTemp" && view.Title != 'Relink Documents');
                    self.getMainRootLibraryViewsSub = new Subscription();
                    self.getMainRootLibraryViews$ = new Subject<any>();
                    self.getMainRootLibraryViews(connector, entityConfiguration, "fields")
                    self.getMainRootLibraryViewsSub = this.getMainRootLibraryViews$.subscribe((getFieldsResponse) => {
                        self.colFields = [];
                        field = getFieldsResponse.d.results;
                        self.colFields = field;
                        selectedView = self.a2dAppService.Views.find((view) => view.DefaultView == true);
                        //let isDefault = selectedView.Title == 'All Documents' ? true : false;
                        viewXml = self.updateView(selectedView.ListViewXml, field, false);
                        self.createViewXml$.next(viewXml);

                    });
                    // self.getFilesAngular(connector, entityConfiguration, path);
                    // self.getFilesAngularSub = this.getFilesAngular$.subscribe((getFilesResponse: any) => {
                    //     field = getFilesResponse.ListSchema.Field;
                    //     selectedView = self.a2dAppService.Views.find((view) => view.DefaultView == true);
                    //     let isDefault = selectedView.Title == 'All Documents' ? true : false;
                    //     viewXml = self.updateView(selectedView.ListViewXml, field, isDefault);
                    //     self.createViewXml$.next(viewXml);
                    // })
                })
            }
            else if (!this.a2dAppService.isValid(this.gridService.selectedViewThumbnail)) {
                //this.getFilesAngular(connector, entityConfiguration, path);
                // this.getFilesAngularSub = this.getFilesAngular$.subscribe((getFilesResponse: any) => {
                //     field = getFilesResponse.ListSchema.Field;
                //     if (this.a2dAppService.isValid(viewId)) {
                //         selectedView = self.a2dAppService.Views.find((view) => view.Id == viewId);
                //     }
                //     else {
                //         selectedView = self.a2dAppService.Views.find((view) => view.DefaultView == true);
                //     }
                //     let isDefault = selectedView.Title == 'All Documents' ? true : false;
                //     viewXml = self.updateView(selectedView.ListViewXml, field, isDefault);
                //     self.createViewXml$.next(viewXml);

                // })
                this.getMainRootLibraryViewsSub = new Subscription();
                this.getMainRootLibraryViews$ = new Subject<any>();
                this.getMainRootLibraryViews(connector, entityConfiguration, "fields")
                this.getMainRootLibraryViewsSub = this.getMainRootLibraryViews$.subscribe((getFieldsResponse) => {
                    self.colFields = [];
                    field = getFieldsResponse.d.results;
                    self.colFields = field;
                    if (this.a2dAppService.isValid(viewId)) {
                        selectedView = self.a2dAppService.Views.find((view) => view.Id == viewId);
                    }
                    else {
                        selectedView = self.a2dAppService.Views.find((view) => view.DefaultView == true);
                    }
                    //let isDefault = selectedView.Title == 'All Documents' ? true : false;
                    viewXml = self.updateView(selectedView.ListViewXml, field, false);
                    self.createViewXml$.next(viewXml);

                });
            }

            else if (this.a2dAppService.isValid(this.gridService.selectedViewThumbnail)) {
                // this.getFilesAngular(connector, entityConfiguration, path);
                // this.getFilesAngularSub = this.getFilesAngular$.subscribe((getFilesResponse: any) => {
                //     field = getFilesResponse.ListSchema.Field;
                //     selectedView = self.a2dAppService.Views.find((view) => view.Id == this.gridService.selectedViewThumbnail);
                //     let isDefault = selectedView.Title == 'All Documents' ? true : false;
                //     viewXml = self.updateView(selectedView.ListViewXml, field, isDefault);
                //     self.createViewXml$.next(viewXml);

                // })
                this.getMainRootLibraryViewsSub = new Subscription();
                this.getMainRootLibraryViews$ = new Subject<any>();
                this.getMainRootLibraryViews(connector, entityConfiguration, "fields")
                this.getMainRootLibraryViewsSub = this.getMainRootLibraryViews$.subscribe((getFieldsResponse) => {
                    self.colFields = [];
                    field = getFieldsResponse.d.results;
                    self.colFields = field;
                    selectedView = self.a2dAppService.Views.find((view) => view.Id == this.gridService.selectedViewThumbnail);
                    //let isDefault = selectedView.Title == 'All Documents' ? true : false;
                    viewXml = self.updateView(selectedView.ListViewXml, field, false);
                    self.createViewXml$.next(viewXml);

                });

            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
            return "";
        }
    }

    updateView(xmlString, jsonObject, bool) {
        let parser = new DOMParser();
        let xmlDoc = parser.parseFromString(xmlString, "text/xml");
        // Check if Query tag should be removed
        // let queryNode = xmlDoc.getElementsByTagName("Query")[0];
        // if (this.a2dAppService.isValid(queryNode)) {
        //     if (this.a2dAppService.isValid(queryNode.getElementsByTagName("FieldRef")[0])) {
        //         let queryFieldRef = queryNode.getElementsByTagName("FieldRef")[0].getAttribute("Name");
        //         let isFieldRefPresent = jsonObject.some(field => field.InternalName === queryFieldRef);
        //         if (!isFieldRefPresent) {
        //             // Remove Query tag if the queryFieldRef is not present in jsonObject
        //             let viewNode = xmlDoc.getElementsByTagName("View")[0];
        //             viewNode.removeChild(queryNode);
        //         }
        //     }
        // }

        let queryNode = xmlDoc.getElementsByTagName("Query")[0];
        let removequery = false;
        if (this.a2dAppService.isValid(queryNode)) {
            let fieldRefNodes = Array.from(queryNode.getElementsByTagName("FieldRef"));
            fieldRefNodes.forEach(fieldRefNode => {
                let fieldName = fieldRefNode.getAttribute("Name");
                if (!jsonObject.some(field => field.InternalName === fieldName)) {
                    removequery = true;
                }
            });
        }

        if (removequery) {
            let viewNode = xmlDoc.getElementsByTagName("View")[0];
            viewNode.removeChild(queryNode);
        }

        // Update ViewFields
        let viewFieldsNodes = xmlDoc.getElementsByTagName("ViewFields")[0];
        if (this.a2dAppService.isValid(viewFieldsNodes)) {
            let viewFieldsArray = Array.from(viewFieldsNodes.children).map(node => node.getAttribute("Name"));
            let realFieldNamesArray = jsonObject.map((field) => field.InternalName);

            // Replace "LinkFilename" with "FileLeafRef" in ViewFields
            // Array.from(viewFieldsNodes.children).forEach(node => {
            //     let fieldName = node.getAttribute("Name");
            //     if (fieldName === "LinkFilename") {
            //         node.setAttribute("Name", "FileLeafRef");
            //     }
            // });

            // Remove extra ViewFields not present in jsonObject
            Array.from(viewFieldsNodes.children).forEach(node => {
                let fieldName = node.getAttribute("Name");
                if (!realFieldNamesArray.includes(fieldName)) {
                    viewFieldsNodes.removeChild(node);
                }
            });

            // Add missing ViewFields from jsonObject if bool is true
            // if (bool) {
            //     let missingFieldNames = realFieldNamesArray.filter(fieldName => !viewFieldsArray.includes(fieldName));
            //     missingFieldNames.forEach(fieldName => {
            //         let newFieldNode = xmlDoc.createElement("FieldRef");
            //         if (fieldName != "FileLeafRef") {
            //             newFieldNode.setAttribute("Name", fieldName);
            //             viewFieldsNodes.appendChild(newFieldNode);
            //         }
            //     });
            // }
        }


        // Serialize updated XML object back to string
        let serializer = new XMLSerializer();
        let updatedXmlString = serializer.serializeToString(xmlDoc);

        return updatedXmlString;
    }



    updateSavedRowData(connector: Connector, entityConfiguration: EntityConfiguration, path: string, rowData?: any): any {
        let functionName: string = "updateSavedRowData";
        let self = this;
        let folderRelativePath: string;
        let firstCharOfPath: string;
        let selectedView;
        folderRelativePath = path.replace(/'/g, "''");
        firstCharOfPath = connector.absolute_url.charAt(0);
        try {

            this.getFilesAngularSub = new Subscription();
            this.getFilesAngular$ = new Subject<any>();
            if (this.a2dAppService.isValid(this.gridService.selectedView)) {
                selectedView = self.a2dAppService.Views.find(view => view.Id == this.gridService.selectedView)
            }
            else {
                selectedView = self.a2dAppService.Views.find(view => view.DefaultView == true);
            }
            let viewXml = this.updateView(selectedView.ListViewXml, this.colFields, false)
            this.getFilesAngular(connector, entityConfiguration, path, "", viewXml, functionName, rowData);
            this.getFilesAngularSub = this.getFilesAngular$.subscribe((getFilesResponse: any) => {
                // console.log("Row Data : " + getFilesResponse);
                let files: any = self.createCollectionSharePointData(getFilesResponse.ListData.Row, self.a2dAppService.columnsArray, connector, entityConfiguration);
                const index = self.utilityService.gridData.findIndex(row => row.UniqueId == files[0].UniqueId);
                if (index != -1) {
                    self.utilityService.gridData[index] = files[0];
                }
                if (self.a2dAppService.isValid(self.gridService.table) && self.a2dAppService.isValid(self.gridService.table.filteredValue)) { // 
                    const index = self.gridService.table.filteredValue.findIndex(item => item.UniqueId === files[0].UniqueId);
                    if (index !== -1) {
                        // Update the properties of the item
                        self.gridService.table.filteredValue[index] = files[0];
                    }
                }
            });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * methods to upload the metdata of the file to SharePoint columns
     * @param connector 
     * @param entityConfiguration 
     * @param UniqueId 
     * @param metaData 
     * @param path 
     */
    uploadMetadataToSharePoint(connector: Connector, entityConfiguration: EntityConfiguration, itemID: string, metaData: any, path: string, viewId?: string, rowData?: any, onSaveFunction?: string) {
        let functionName: string = "uploadMetadataToSharePoint";
        let self = this;
        let httpOptions = {};
        let sharePointSite: string;
        let requestUrl: string;
        let data: string;
        let rootFolderName: string;
        let rootFolderUrl: string;
        try {
            //@ts-ignore
            this.acceessToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;

            sharePointSite = connector.absolute_url;

            rootFolderName = entityConfiguration.folder_path.replace(/^\/+|\/+$/g, '').split('/')[0];

            if (sharePointSite.split('/').length > 3) {
                let sitePath: any = sharePointSite.split('/').splice(3).join('/');
                rootFolderUrl = '/' + sitePath + "/" + rootFolderName;
            }
            else {
                rootFolderUrl = rootFolderName;
            }
            //requestUrl = sharePointSite + "/_api/web/getFileById('" + UniqueId + "')/ListItemAllFields";
            requestUrl = sharePointSite + "/_api/web/GetList('" + rootFolderUrl + "')/items('" + itemID + "')/ValidateUpdateListItem()";


            // Assuming metaData is available in the context and populated
            let jsonData = { formValues: [] };

            if (onSaveFunction != "createSharePointMetadata") {
                this.gridService.cols.forEach((column) => {
                    const field = column.field;
                    if (metaData.hasOwnProperty(field) && field != "Modified") {
                        let fieldValue: any = metaData[field] != null ? metaData[field].toString() : ""; // Use 'any' type for generalization

                        // Process the data based on its type
                        // if (column.fieldType == "datetime" || column.fieldType == "dateonly") {
                        //     // Convert the date to ISO 8601 string without milliseconds
                        //     fieldValue = new Date(fieldValue).toISOString().slice(0, 16);
                        // } else
                        if (column.fieldType == "boolean") {
                            // Convert "Yes"/"No" to true/false
                            fieldValue = (fieldValue == "Yes") ? "1" : "0";
                        }

                        // Append the processed data in the required format
                        jsonData.formValues.push({
                            FieldName: field,
                            FieldValue: fieldValue
                        });
                    }
                });
            }
            else {
                Object.keys(metaData).forEach((key) => {
                    let fieldValue: any;
                    // Access the inner object corresponding to the key
                    const innerObject = metaData[key];
                    // Access the type property of the inner object
                    const type = innerObject.type;
                    const value = innerObject.value;
                    console.log(`Type of ${key}: ${type}`);
                    if (type == "datetime" && this.a2dAppService.isValid(value) || type == "dateonly" && this.a2dAppService.isValid(value)) {
                        // Convert the date to ISO 8601 string without milliseconds
                        fieldValue = this.utilityService.getDateTimeInUserTZ(this.a2dAppService.crmUserTimeZoneParameter["TimeZoneBias"], value, this.a2dAppService.crmUserTimeZoneParameter["dateformatstring"], this.a2dAppService.crmUserTimeZoneParameter["timeformatstring"]);
                    } else
                        if (type == "boolean" && this.a2dAppService.isValid(value)) {
                            // Convert "Yes"/"No" to true/false
                            fieldValue = (value == "Yes") ? "1" : "0";
                        } else {
                            fieldValue = value;
                        }

                    if (this.a2dAppService.isValid(fieldValue)) {
                        jsonData.formValues.push({
                            FieldName: key,
                            FieldValue: fieldValue
                        });
                    }
                    // Append the processed data in the required format

                });

            }

            data = JSON.stringify(jsonData);
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + this.acceessToken,
                    "Content-Type": "application/json;odata=verbose",
                    "If-Match": "*"
                }),
            };
            this.executePost(requestUrl, data, httpOptions, connector, entityConfiguration).then(response => {
                if (response.hasError == true && typeof response.comments == 'string') {
                    if ('comments' in response) {
                        if (response.comments.includes("2130575257") || response.comments.includes("tokenExpired")) {
                            // console.log("Response of " + functionName + ": ", response);
                            //this.getSharePointData(connector, entityConfiguration, metaData["path_display"]);
                            self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then(
                                (response: any) => {
                                    this.acceessToken = self.a2dAppService.isValid(response) ? response.access_token : null;
                                    self.uploadMetadataToSharePoint(connector, entityConfiguration, itemID, metaData, path, viewId, rowData, onSaveFunction);

                                },
                                (error: any) => {
                                    self.getFilesAngular$.next(error);
                                }

                            );
                        }
                        else {
                            if (self.modalService.isOpen == false) {
                                if (response.message) {
                                    self.a2dAppService.retrieveEntityDefinitions('', response.message.value + ' - { From uploadMetadataToSharePoint }', entityConfiguration, '', '');
                                    self.a2dAppService.entityDefinitionSub = self.a2dAppService.entityDefinitions$.subscribe(
                                        (response) => {
                                            self.a2dAppService.objectTypeCodeArray[self.a2dAppService.currentEntityName.toLowerCase()] = response[0].ObjectTypeCode;
                                            self.createErrorLog(response["errorResponse"], entityConfiguration, response[0].ObjectTypeCode, path);
                                        }
                                    );
                                }
                            }
                            if (onSaveFunction == "createSharePointMetadata") {
                                self.createSharePointMetadata$.next(true);
                                // this.getSharePointData(connector, entityConfiguration, path, viewId, rowData, onSaveFunction);
                            }
                            else {
                                this.updateSavedRowData(connector, entityConfiguration, path, rowData);
                            }
                        }
                    }
                    else {
                        if (self.modalService.isOpen == false) {
                            if (response.message) {
                                self.a2dAppService.retrieveEntityDefinitions('', response.message.value + ' - { From uploadMetadataToSharePoint }', entityConfiguration, '', '');
                                self.a2dAppService.entityDefinitionSub = self.a2dAppService.entityDefinitions$.subscribe(
                                    (response) => {
                                        self.a2dAppService.objectTypeCodeArray[self.a2dAppService.currentEntityName.toLowerCase()] = response[0].ObjectTypeCode;
                                        self.createErrorLog(response["errorResponse"], entityConfiguration, response[0].ObjectTypeCode, path);
                                    }
                                );
                            }
                        }                         //this.getSharePointData(connector, entityConfiguration, path, viewId, rowData, onSaveFunction);
                        if (onSaveFunction == "createSharePointMetadata") {
                            self.createSharePointMetadata$.next(true);

                            // this.getSharePointData(connector, entityConfiguration, path, viewId, rowData, onSaveFunction);

                        }
                        else {
                            this.updateSavedRowData(connector, entityConfiguration, path, rowData);
                        }
                    }
                }
                else {
                    if (self.modalService.isOpen == false) {
                        if (response.message) {
                            self.a2dAppService.retrieveEntityDefinitions('', response.message.value + ' - { From uploadMetadataToSharePoint }', entityConfiguration, '', '');
                            self.a2dAppService.entityDefinitionSub = self.a2dAppService.entityDefinitions$.subscribe(
                                (response) => {
                                    self.a2dAppService.objectTypeCodeArray[self.a2dAppService.currentEntityName.toLowerCase()] = response[0].ObjectTypeCode;
                                    self.createErrorLog(response["errorResponse"], entityConfiguration, response[0].ObjectTypeCode, path);
                                }
                            );
                        }
                    }                     //this.getSharePointData(connector, entityConfiguration, path, viewId, rowData, onSaveFunction);
                    if (onSaveFunction == "createSharePointMetadata") {
                        self.createSharePointMetadata$.next(true);
                        // this.getSharePointData(connector, entityConfiguration, path, viewId, rowData, onSaveFunction);

                    }
                    else {
                        this.updateSavedRowData(connector, entityConfiguration, path, rowData);
                    }
                }
            }
            ).catch((err: any) => {
                self.utilityService.throwError(err, functionName);
            });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    /**
     * Method to create the key value pairs of SHarePoint Column and their values to upload the metadata to SharePoint
     * @param connector 
     * @param entityConfiguration 
     * @param UniqueId 
     * @param path 
     */
    createSharePointMetadata(connector: Connector, entityConfiguration: EntityConfiguration, itemID: string, path: string) {
        let functionName: string = "createSharePointMetadata";
        let metaDataJSONArray: any = [];
        let tempArray: any = [];
        let self = this;
        let lookUpPersonCounter: any;
        let successCounter: any = 0;
        let matchingRecord: any;
        //const fieldValueMap: { [key: string]: any } = {};
        const fieldValueMap: { [key: string]: { type: string, value: any } } = {};
        try {
            if ((entityConfiguration.isActivity && entityConfiguration.activityMetadataEnabled) || entityConfiguration.linearMetadataEnabled) {

                if (this.a2dAppService.isValid(this.a2dAppService._recordValuesArray)) {

                    if (this.a2dAppService.isValid(entityConfiguration.metadataJSON) && (this.a2dAppService.isValid(entityConfiguration.parentHierarchyRecordMetadaJSON))) {
                        metaDataJSONArray.push(JSON.parse(entityConfiguration.metadataJSON));
                        metaDataJSONArray.push(JSON.parse(entityConfiguration.parentHierarchyRecordMetadaJSON));
                        if (this.a2dAppService.isActivity && entityConfiguration.isActivityFolderCreation && (this.a2dAppService.isValid(entityConfiguration.activityMetadaJSON))) {
                            metaDataJSONArray.push(JSON.parse(entityConfiguration.activityMetadaJSON));
                        }
                    }
                    else {
                        if (this.a2dAppService.isValid(entityConfiguration.metadataJSON)) {
                            metaDataJSONArray.push(JSON.parse(entityConfiguration.metadataJSON));
                        }
                    }
                    if (this.a2dAppService.isValid(metaDataJSONArray) && metaDataJSONArray.length > 0) {

                        tempArray = this.a2dAppService._recordValuesArray
                            .map(record => metaDataJSONArray.find(metaData => record.entitySetName === metaData.entitySetName))
                            .filter(metaData => metaData !== undefined);

                        let currenEntityName = this.a2dAppService.isValid(entityConfiguration.activityLogicalName) ? entityConfiguration.activityLogicalName : this.a2dAppService.currentEntityName;
                        let updatedData = this.removeDuplicatesBasedOnEntityNameCondition(tempArray, currenEntityName);
                        lookUpPersonCounter = updatedData.flatMap(item => item.fields).filter(a => a.fieldType == "lookup" || a.fieldType === 'user').length
                        this.a2dAppService._recordValuesArray.map((record: any) => {
                            updatedData.map((crmEntity: CRMEntity) => {
                                const entityName = crmEntity.entityName;
                                crmEntity.fields.map((field: Field) => {
                                    const sharePointColumn = field.sharePointColumn;
                                    const sharePointColumnTypeAsString = field.sharePointColumnTypeAsString.toLowerCase();
                                    const crmField = field.fieldLogicalName;

                                    if ((record.entitySetName == crmEntity.entitySetName)) {
                                        let formattedValue = "";
                                        //Instead of the nested if else used Switch case
                                        switch (sharePointColumnTypeAsString) {
                                            case "choice":
                                                formattedValue = this.a2dAppService.isValid(record[0][crmField + "@OData.Community.Display.V1.FormattedValue"]) ? record[0][crmField + "@OData.Community.Display.V1.FormattedValue"].toString() : "";
                                                fieldValueMap[sharePointColumn] = { type: sharePointColumnTypeAsString, value: formattedValue };
                                                break;
                                            case "lookup":
                                                matchingRecord = this.a2dAppService._recordValuesArray.find(record => record.entityName === entityName.toLowerCase());
                                                let recordId: string = this.a2dAppService.isValid(matchingRecord[0]) && this.a2dAppService.isValid(matchingRecord[0][`_${field.fieldLogicalName}_value`]) ? matchingRecord[0][`_${field.fieldLogicalName}_value`] : "";

                                                let displayName: string = this.a2dAppService.isValid(matchingRecord[0]) && this.a2dAppService.isValid(matchingRecord[0][`_${field.fieldLogicalName}_value@OData.Community.Display.V1.FormattedValue`]) ? matchingRecord[0][`_${field.fieldLogicalName}_value@OData.Community.Display.V1.FormattedValue`] : "";
                                                let lookUpEntityName: string = this.a2dAppService.isValid(matchingRecord[0]) && this.a2dAppService.isValid(matchingRecord[0][`_${field.fieldLogicalName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`]) ? matchingRecord[0][`_${field.fieldLogicalName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`] : "";
                                                if (self.a2dAppService.isValid(recordId) && self.a2dAppService.isValid(lookUpEntityName)) {
                                                    self.App_AddLookUpListItems$ = new Subject<any>();
                                                    self.App_AddLookUpListItems(connector, entityConfiguration, recordId, displayName, lookUpEntityName, field.sharePointColumnListId, sharePointColumn, true);
                                                    self.App_AddLookUpListItemsSub = self.App_AddLookUpListItems$.subscribe((response: any) => {
                                                        successCounter++;
                                                        if (self.a2dAppService.isValid(response)) {
                                                            formattedValue = this.a2dAppService.isValid(response) ? response.toString() : "";
                                                            fieldValueMap[response.key] = { type: 'lookup', value: response.value };
                                                            if (successCounter == lookUpPersonCounter) {
                                                                this.uploadMetadataToSharePoint(connector, entityConfiguration, itemID, fieldValueMap, path, this.gridService.selectedView, null, functionName);
                                                            }
                                                        }
                                                    })
                                                }
                                                else {
                                                    successCounter++;
                                                    if (successCounter == lookUpPersonCounter) {
                                                        this.uploadMetadataToSharePoint(connector, entityConfiguration, itemID, fieldValueMap, path, this.gridService.selectedView, null, functionName);
                                                    }
                                                }
                                                break;
                                            case "user":

                                                matchingRecord = this.a2dAppService._recordValuesArray.find(record => record.entityName === entityName.toLowerCase());
                                                let Id: string = this.a2dAppService.isValid(matchingRecord[0]) && this.a2dAppService.isValid(matchingRecord[0][`_${field.fieldLogicalName}_value`]) ? matchingRecord[0][`_${field.fieldLogicalName}_value`] : "";
                                                if (this.a2dAppService.isValid(Id)) {
                                                    self.App_AddUserListItems$ = new Subject<any>();
                                                    self.App_AddUserListItems(connector, entityConfiguration, Id, field.fieldLogicalName, sharePointColumn)
                                                    // Get the Email id of the used by using the GUID of the systemUser
                                                    self.App_AddUserListItemsSub = self.App_AddUserListItems$.subscribe((response: any) => {
                                                        successCounter++;
                                                        formattedValue = this.a2dAppService.isValid(response) ? response.toString() : "";

                                                        fieldValueMap[response.key] = { type: 'user', value: "[{\"key\":\"i:0#.f|membership|" + response.value + "\"}]" };
                                                        if (successCounter == lookUpPersonCounter) {
                                                            this.uploadMetadataToSharePoint(connector, entityConfiguration, itemID, fieldValueMap, path, this.gridService.selectedView, null, functionName);
                                                        }
                                                    });
                                                }
                                                else {
                                                    successCounter++;
                                                    if (successCounter == lookUpPersonCounter) {
                                                        this.uploadMetadataToSharePoint(connector, entityConfiguration, itemID, fieldValueMap, path, this.gridService.selectedView, null, functionName);
                                                    }
                                                }
                                                break;
                                            default:
                                                formattedValue = this.a2dAppService.isValid(record[0][crmField]) ? record[0][crmField].toString() : "";
                                                fieldValueMap[sharePointColumn] = { type: sharePointColumnTypeAsString, value: formattedValue };
                                                break;

                                        }
                                        // if (sharePointColumnTypeAsString == "choice") {

                                        // }
                                        // else if(sharePointColumnTypeAsString == "lookup")
                                        //     {

                                        //     }
                                        //  else {
                                        //   }
                                        //fieldValueMap[sharePointColumn] = { type: sharePointColumnTypeAsString, value: formattedValue };
                                    }
                                });
                            });
                        });

                        if (this.a2dAppService.isValid(fieldValueMap) && successCounter == lookUpPersonCounter) {
                            this.uploadMetadataToSharePoint(connector, entityConfiguration, itemID, fieldValueMap, path, this.gridService.selectedView, null, functionName);
                        }
                    }
                    else {
                        self.createSharePointMetadata$.next(true);
                    }
                }
                else {
                    self.createSharePointMetadata$.next(true);
                }
            }
            else {
                self.createSharePointMetadata$.next(true);
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    // Find all common sharePointColumnTypeDisplayName properties across all objects
    removeDuplicatesBasedOnEntityNameCondition(arr, currentEntityName) {
        const displayNameCounts = {};

        // Count occurrences of each sharePointColumnTypeDisplayName
        arr.forEach(obj => {
            obj.fields.forEach(field => {
                const displayName = field.sharePointColumn;
                if (!displayNameCounts[displayName]) {
                    displayNameCounts[displayName] = 0;
                }
                displayNameCounts[displayName]++;
            });
        });

        // Remove fields with duplicate sharePointColumnTypeDisplayName based on the entity name condition
        return arr.map(obj => {
            if (obj.entityName !== currentEntityName) {
                obj.fields = obj.fields.filter(field => displayNameCounts[field.sharePointColumn] <= 1);
            }
            return obj;
        });
    }

    /**
    * Create collection based on the result
    * @param result this is an collectoion of files
    * return array
    */
    createCollectionSharePointData(result: any, columns: any, connector: Connector, entityConfiguration: EntityConfiguration): any {
        let GridDataList: GridData[] = [];
        let functionName: string = "createCollectionSharePointData";
        try {
            for (let index = 0; index < result.length; index++) {
                let data: any = {};
                let element: any = result[index];
                columns.forEach((col: any) => {
                    //#region 
                    if (col.RealFieldName != "DocIcon") {
                        switch (col.FieldType) {
                            case "Number":
                            case "Currency":
                                data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName])) ? { label: element[col.RealFieldName], value: element[col.RealFieldName + "."].split(".")[0] } : { label: "", value: "" };
                                break;
                            case "DateTime":
                                data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName])) ? { label: element[col.RealFieldName], value: new Date(element[col.RealFieldName]) } : { label: "", value: null };
                                break;
                            case "Lookup":
                                data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName]) && this.a2dAppService.isValid(element[col.RealFieldName][0])) ? { label: this.a2dAppService.isValid(element[col.RealFieldName][0].lookupValue) ? element[col.RealFieldName][0].lookupValue : element[col.RealFieldName], value: this.a2dAppService.isValid(element[col.RealFieldName][0].lookupId) ? element[col.RealFieldName][0].lookupId : element[col.RealFieldName] } : { label: "", value: "" };
                                break;
                            case "User":
                                data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName]) && this.a2dAppService.isValid(element[col.RealFieldName][0])) ? { label: element[col.RealFieldName][0].title, value: element[col.RealFieldName][0].id } : { label: "", value: "" };
                                break;
                            default:
                                data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName])) ? element[col.RealFieldName] : "";
                                break;
                        }
                    }
                    //#endregion
                    // if (col.FieldType != "User" && col.RealFieldName != "DocIcon" && col.FieldType != "Number" && col.FieldType != "Currency" && col.FieldType != "DateTime") {
                    //     data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName])) ? element[col.RealFieldName] : "";
                    // }
                    // if (col.FieldType != "User" && col.RealFieldName != "DocIcon" && (col.FieldType == "Number" || col.FieldType == "Currency")) {
                    //     data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName])) ? { label: element[col.RealFieldName], value: element[col.RealFieldName + "."].split(".")[0] } : { label: "", value: "" };
                    // }
                    // if (col.FieldType != "User" && col.RealFieldName != "DocIcon" && col.FieldType == "DateTime") {
                    //     data[col.RealFieldName] = (this.a2dAppService.isValid(element[col.RealFieldName])) ? { label: element[col.RealFieldName], value: new Date(element[col.RealFieldName]) } : { label: "", value: null };
                    // }
                })
                data.ID = this.a2dAppService.isValid(element.ID) ? element.ID : "";
                data.fileType = this.a2dAppService.isValid(element.FSObjType) ? element.FSObjType != 0 ? "folder" : "file" : "";
                data.UniqueId = this.a2dAppService.isValid(element.UniqueId) ? element.UniqueId.replace('{', '').replace('}', '') : "";
                data.fieldType = this.a2dAppService.isValid(element.FieldType) ? element.FieldType : "";
                data.path_display = this.a2dAppService.isValid(element.FileRef) ? element.FileRef : "";
                data.isChecked = false;
                //Added by Lakshman for bulk data update isEditEnabled and index
                data.isEditEnabled = true;
                data.index = index;
                data.isEditActive = false;

                if (this.a2dAppService.selectedEntityRecords > 0) {
                    data.size = this.a2dAppService.isValid(element.size) ? this.Convert((parseInt(element.size) / 1024).toString()) : "";

                }
                else {
                    data.File_x0020_Size = this.a2dAppService.isValid(element.File_x0020_Size) ? { label: this.Convert((parseInt(element.File_x0020_Size) / 1024).toString()), value: (parseInt(element.File_x0020_Size) / 1024) } : { label: "", value: "" };
                }
                data.fileUrl = this.utilityService.getSPThumbnails(connector, entityConfiguration, element);
                GridDataList.push(data);
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return GridDataList;
    }

    createErrorLog(errorResponse: any, selectedEntityConfiguration: any, objectTypeCode: any, path?: string): void {
        let clientURL: any = null;
        let recordURL: any = null;
        let functionName: string = "createErrorLog";
        let self: any = this;
        let file: any = {};
        let curRecordId: any = "";
        try {
            // let currentdate = new Date();
            // let userId = this._Xrm.getUserId().substring(1, this._Xrm.getUserId().length - 1);
            clientURL = this.a2dAppService._Xrm.getClientUrl();
            curRecordId = this.a2dAppService.currentEntityId;
            recordURL = clientURL + "/main.aspx?etc=" + objectTypeCode + "&id=%7b" + this.a2dAppService.currentEntityId + "%7d&pagetype=entityrecord";
            if (this.a2dAppService.isValid(errorResponse)) {
                //create Error log data object 
                file = {
                    "ikl_recordid": curRecordId,
                    "ikl_EntityConfiguration@odata.bind": "/ikl_entityconfigurations(" + selectedEntityConfiguration.entity_configurationid + ")",
                    "ikl_error": errorResponse,
                    "ikl_recordurl": recordURL,
                    "ikl_filepath": this.a2dAppService.isValid(path) ? path : ""
                };
            }
            this.a2dAppService.webApi.create("ikl_errorlogs", file).then((response: any) => {
                response = this.a2dAppService.extractResponse(response);
            }, (error: any) => {
                this.a2dAppService.throwError(error, functionName);
            });
        } catch (error) {
            this.a2dAppService.throwError(error, functionName);
        }
    }

    /**
     * Create Folder and Upload Files entry point
     * @param folderPathArray
     * @param uploadPath
     * @param workItems
     * @param selectedConnectorTab
     * @param selectedEntityConfiguration
     */
    createFolderAndUploadFiles(folderPathArray: string[], uploadPath: string, workItems: any, selectedConnectorTab: Connector, selectedEntityConfiguration: EntityConfiguration, value: any): void {
        let functionName: string = "createFolderAndUploadFiles";
        let self = this;
        try {
            let folderPathCollection: {} = {};
            folderPathCollection["Folders"] = folderPathArray;
            folderPathCollection["UploadPath"] = uploadPath;
            this.createFolders(folderPathCollection, selectedConnectorTab, selectedEntityConfiguration);
            this.createFolderAndUploadFilesSub = this.createFolders$.subscribe(
                (response) => {
                    if (workItems.length > 0) {
                        self.uploadFileToSP(workItems, uploadPath, selectedConnectorTab, selectedEntityConfiguration, "UploadFolder");
                    }
                    else {
                        self.spinnerService.hide();
                        if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                            self.modalService.openDialogWithInputUploadStatus(self.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onOKClick) => {
                            });
                        }
                        else {
                            self.getSharePointData(selectedConnectorTab, selectedEntityConfiguration, uploadPath, self.gridService.selectedView);
                        }
                    }
                    self.createFolderAndUploadFilesSub.unsubscribe();
                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        self.spinnerService.hide();
                    }
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
       * Create Folder and Upload Files entry point
       * @param folderPathArray
       * @param uploadPath
       * @param workItems
       * @param selectedConnectorTab
       * @param selectedEntityConfiguration
       */
    createFolderAndUploadFilesHome(folderPathArray: string[], uploadPath: string, workItems: any, selectedConnectorTab: Connector, selectedEntityConfiguration: EntityConfiguration, value: any): void {
        let functionName: string = "createFolderAndUploadFilesHome";
        let self = this;
        try {
            let folderPathCollection: {} = {};
            folderPathCollection["Folders"] = folderPathArray;
            folderPathCollection["UploadPath"] = uploadPath;
            this.createFolders(folderPathCollection, selectedConnectorTab, selectedEntityConfiguration);
            this.createFolderAndUploadFilesSub = this.createFolders$.subscribe(
                (response) => {
                    self.createFolderAndUploadFilesSub.unsubscribe();
                    if (workItems.length > 0) {
                        self.uploadFileToSPHomeGrid(workItems, uploadPath, selectedConnectorTab, response["entity"], "UploadFolder", value, self.a2dAppService.selectedEntityRecords.length);
                    }
                    else {
                        self.spinnerService.hide();
                        if (this.a2dAppService.selectedEntityRecords.length == 0) {
                            if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                                self.modalService.openDialogWithInputUploadStatus(self.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onOKClick) => {
                                });
                            }
                            else {
                            }
                        }
                    }
                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        self.spinnerService.hide();
                    }
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    parseMetadataJson(jsonString: any): any {
        return JSON.parse(jsonString)
    }

    /**
     * Business Logic for uploading file
     * @param workItems
     * @param uploadPath
     * @param selectedConnectorTab
     * @param selectedEntityConfiguration
     * @param source
     */
    async uploadFileToSP(workItems: any, uploadPath: string, selectedConnectorTab: Connector, selectedEntityConfiguration: EntityConfiguration, source: string) {
        let functionName: string = "uploadFileToSP";
        let additionalWaitTime: number = 1;
        let standardWaitTime: number = 1000;
        try {

            let self = this;

            let count = workItems.length;
            //Shreyas : 18b April 2022.
            let runningCount = 1;
            //Remove the subsite component from the Path
            let subSite: string = this.utilityService.getSharePointSubSite(selectedConnectorTab.absolute_url);
            uploadPath = this.utilityService.clearSubSiteFromPath(subSite, uploadPath);

            let fileEntityConfigurationID: any = "";
            if (self.a2dAppService.isSharePointSecuritySyncLicensePresent && selectedConnectorTab.isSyncRecords) {
                fileEntityConfigurationID = await self.getIklFileEntityConfiguration(selectedConnectorTab);
            }

            self.a2dAppService.UploadingFileCount = 0;
            self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['uploadingfile'] + " " + self.a2dAppService.UploadingFileCount + "/" + (workItems.length + self.a2dAppService.IgnoreFileCount);
            if (workItems.length > 0) {
                self.modalService.fileDownloadingPercentage = "";
                self.modalService.fileUploadingPercentage = "0% Uploaded";
                self.modalService.openUploadStatus(self.modalService.displayMessage, (onClose) => {
                });
                additionalWaitTime = self.getAdditionalTimeValue(count);
                for (let i = 0; i < workItems.length + this.a2dAppService.IgnoreFileCount; i++) {
                    if (this.a2dAppService.isValid(workItems[i])) {
                        try {
                            let decryptedToken: any = "";
                            let workItem = workItems[i];
                            let file = workItem.file;
                            let path = this.utilityService.formatNameWithOutSlash(workItem.path);
                            let name = this.utilityService.formatFileName(workItem.file.name, selectedConnectorTab.connector_type_value);
                            //let name = workItem.file.name;
                            path = source == "UploadFolder" ? `${uploadPath}${path}` : uploadPath;
                            //self.UploadFiles(name, path, base64, selectedConnectorTab, selectedEntityConfiguration);
                            //Shrujan 13 feb 22 Added new method to upload files below 250 Mb size.
                            //self.uploadSPFiles(name, path, base64, selectedConnectorTab, selectedEntityConfiguration, decryptedToken);//shrujan
                            //Shrujan 09 Aug 23 Added new method to upload files below 1.5 GB size.
                            const uploadResp: any = await self.uploadFilesSPmain(file, path, name, selectedConnectorTab, selectedEntityConfiguration, source);
                            if (self.a2dAppService.isValid(uploadResp)) {
                                if (uploadResp.status == "true") {
                                    self.a2dAppService.SuccessFileCount = self.a2dAppService.SuccessFileCount + 1;
                                    self.a2dAppService.UploadingFileCount = self.a2dAppService.UploadingFileCount + 1;
                                    self.modalService.fileUploadingPercentage = "0% Uploaded";
                                    self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['uploadingfile'] + " " + self.a2dAppService.UploadingFileCount + "/" + (workItems.length + self.a2dAppService.IgnoreFileCount);
                                    self.a2dAppService.SuccessFileNames.push(self.createUploadedFileDetailsObject(uploadResp));

                                    if (self.a2dAppService.isValid(fileEntityConfigurationID) && this.isIkl_FilePrivillagesValid) {
                                        let fileLogicalName = "ikl_file";
                                        self.createSyncStatusForFile(uploadResp.FileName, uploadResp.FilePath, uploadResp.FileUniqueId, fileLogicalName, fileEntityConfigurationID, selectedConnectorTab, selectedEntityConfiguration)
                                    }

                                }
                                else if (uploadResp.status == "false") {
                                    self.a2dAppService.ErrorFileCount = self.a2dAppService.ErrorFileCount + 1;
                                    self.a2dAppService.UploadingFileCount = self.a2dAppService.UploadingFileCount + 1;
                                    self.modalService.fileUploadingPercentage = "0% Uploaded";
                                    self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['uploadingfile'] + " " + self.a2dAppService.UploadingFileCount + "/" + (workItems.length + self.a2dAppService.IgnoreFileCount);
                                    self.a2dAppService.ErrorFileNames.push(self.createUploadedFileDetailsObject(uploadResp));
                                    //self.a2dAppService.logError(workItem, uploadResp.description || uploadResp.message, selectedEntityConfiguration, uploadResp.counter, uploadResp["FileName"], uploadResp["FilePath"]);

                                }
                                if (runningCount == count) {
                                    self.spinnerService.hide();
                                    if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                                        self.modalService.UploadStatusModalRef.hide();
                                        self.modalService.openDialogWithInputUploadStatus(self.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onClose) => {
                                            self.getSharePointData(selectedConnectorTab, selectedEntityConfiguration, uploadPath, self.gridService.selectedView);
                                        });
                                    }
                                    else {
                                        self.modalService.UploadStatusModalRef.hide();
                                        self.getSharePointData(selectedConnectorTab, selectedEntityConfiguration, uploadPath, self.gridService.selectedView);
                                    }
                                }
                                // Increment the running count
                                runningCount++;
                            }
                            else {
                                if (self.modalService.isOpen == false) {
                                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                }
                            }
                        }
                        catch (error) {
                            if (self.modalService.isOpen == false) {
                                self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                            }
                            // Based on the Running Count and Count, determine the point to turn off the spinner
                            if (runningCount == count) {
                                self.spinnerService.hide();
                                if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                                    self.modalService.UploadStatusModalRef.hide();
                                    self.modalService.openDialogWithInputUploadStatus(self.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onCloseClick) => {
                                        self.getSharePointData(selectedConnectorTab, selectedEntityConfiguration, uploadPath, self.gridService.selectedView);
                                    });
                                }
                                else {
                                    self.modalService.UploadStatusModalRef.hide();
                                    self.getSharePointData(selectedConnectorTab, selectedEntityConfiguration, uploadPath, self.gridService.selectedView);
                                }
                            }
                            // Increment the running count
                            runningCount++;
                            self.utilityService.throwError(error, functionName);
                        }
                    }
                    else {
                        self.a2dAppService.UploadingFileCount = self.a2dAppService.UploadingFileCount + 1;
                        self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['uploadingfile'] + " " + self.a2dAppService.UploadingFileCount + "/" + (workItems.length + self.a2dAppService.IgnoreFileCount);
                    }
                }
            }
            else {
                this.spinnerService.hide();
                if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                    self.modalService.UploadStatusModalRef.hide();
                    this.modalService.openDialogWithInputUploadStatus(this.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onCloseClick) => {
                        this.getSharePointData(selectedConnectorTab, selectedEntityConfiguration, uploadPath, self.gridService.selectedView);
                    });
                }
                else {
                    self.modalService.UploadStatusModalRef.hide();
                    self.getSharePointData(selectedConnectorTab, selectedEntityConfiguration, uploadPath, self.gridService.selectedView);
                }
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }


    //#region UI File syncing 

    async getIklFileEntityConfiguration(connector: Connector): Promise<string> {
        const functionName = "getIklFileEntityConfiguration : ";
        let fetchXML = "";
        let self = this;
        let queryOptions: any = null;
        let extractedResponse: any;
        try {
            queryOptions = {
                includeFormattedValues: true,
                includeLookupLogicalNames: true,
                includeAssociatedNavigationProperties: true
            }

            fetchXML = `<fetch version='1.0' mapping='logical' distinct='true' >
            <entity name='ikl_entityconfiguration' >
                <attribute name='ikl_name' />
                <attribute name='statecode' />
                <attribute name='ikl_entityname' />
                <attribute name='ikl_connector' />
                <attribute name='ikl_entityconfigurationid' />
                <filter type='and' >
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='ikl_entityname' operator='eq' value='ikl_file' />
                </filter>
                <link-entity name='ikl_connector' alias='aa' link-type='inner' from='ikl_connectorid' to='ikl_connector' >
                    <filter type='and' >
                        <condition attribute='ikl_connectorid' operator='eq' value='${connector.connector_id}' />
                    </filter>
                </link-entity>
            </entity>
        </fetch>`;
            fetchXML = encodeURIComponent(fetchXML);

            const response = await self.a2dAppService.webApi.retrieveMultiple("ikl_entityconfigurations", "fetchXml=" + fetchXML, queryOptions);
            extractedResponse = self.a2dAppService.extractResponse(response);
            return (extractedResponse && extractedResponse.length > 0) ? extractedResponse[0].ikl_entityconfigurationid : '';

        } catch (error) {
            console.log(`${functionName} : error message ${error.message}`);
            return '';
        }
    }

    private async createSyncStatusForFile(
        fileName: string,
        absoluteUrl: string,
        fileUniqueId: string,
        fileLogicalName: string,
        fileEntityConfigurationID: any, // Use the appropriate type
        connector: Connector, // Use the appropriate type for Connector
        entityConfiguration: EntityConfiguration
    ): Promise<void> {
        const functionName = "createSyncStatusForNotes";

        let relativeUrl = "";

        let fileActionJsonStr = "";
        let entityRoles: any[] = [];
        let securityPrivileges: any[] = [];
        let connectionDetailsId: string = "";
        let recordEntity: any = null;
        let groupNames: string[] = [];
        let userPermissionId = "";
        let groupCollection: any = null;

        let regardingRef: any = null;
        let userMaxPrivilege: UserPrivilegeData = null;
        let self = this;
        let syncStatusID: any;
        let filemaindocLocId: any;
        let fetchCondition: any;
        let loggedInUser: any;
        try {

            if (self.a2dAppService.isSharePointSecuritySyncLicensePresent) {

                relativeUrl = absoluteUrl + "/" + fileName;

                this.retrieveSyncStatus(fileEntityConfigurationID, connector);

                self.syncStatuesSub = self.syncStatues$.subscribe((syncStatusId) => {
                    self.syncStatuesSub.unsubscribe();
                    if (self.a2dAppService.isValid(syncStatusId)) {

                        syncStatusID = syncStatusId;

                        regardingRef = { logicalName: fileLogicalName, id: "" };

                        this.retrieveSharePointLocation(regardingRef, true, connector.sharepoint_site_id);

                        self.spdoclocSub = self.spdocloc$.subscribe((doclocationresponseId: any) => {
                            self.spdoclocSub.unsubscribe();

                            if (self.a2dAppService.isValid(doclocationresponseId)) {
                                filemaindocLocId = doclocationresponseId;

                                const entRole = { EntityName: fileLogicalName, EntityConfigId: fileEntityConfigurationID, ConnectorId: connector.connector_id };
                                entityRoles.push(entRole);

                                this.retrieveSecurityMetadata(entityRoles, securityPrivileges);
                                self.entitySecurityMetadataFetchSub = self.entitySecurityMetadataFetch$.subscribe((fetchxmlresp: any) => {
                                    self.entitySecurityMetadataFetchSub.unsubscribe();
                                    fetchCondition = fetchxmlresp;

                                    if (self.a2dAppService.isValid(fetchCondition)) {
                                        loggedInUser = self.a2dAppService._Xrm.getUserId();
                                        loggedInUser = loggedInUser.substring(1, loggedInUser.length - 1);
                                        this.getMaxPrivilegeMask(loggedInUser, fetchCondition, connector.connector_id);
                                        self.userMaxPrevSub = self.userMaxPrev$.subscribe((userPrevResp: UserPrivilegeData) => {
                                            self.userMaxPrevSub.unsubscribe();
                                            userMaxPrivilege = userPrevResp;

                                            if (self.a2dAppService.isValid(userMaxPrivilege)) {

                                                groupNames = self.getGroupNames(fileLogicalName, userMaxPrivilege, connector, 0);

                                                const fileActionJson = {
                                                    docLocId: filemaindocLocId,
                                                    entityConfigurationId: fileEntityConfigurationID.toString(),
                                                    entityLogicalName: "spdocloc",
                                                    entityName: relativeUrl,
                                                    parentDocumentLocation: filemaindocLocId,
                                                    parentLocRelativeUrl: fileLogicalName,
                                                    permissionId: userPermissionId || null,
                                                    regardingId: loggedInUser,
                                                    regardingLogicalName: fileLogicalName,
                                                    sharepointSiteId: connector.sharepoint_site_id,
                                                    spUserId: userMaxPrivilege?.spUserId || null,
                                                    supportedEntityLogicalName: "spdocloc",
                                                    fileUniqueId: fileUniqueId
                                                };
                                                fileActionJsonStr = JSON.stringify(fileActionJson);

                                                if (groupNames.includes("user")) {
                                                    self.retrievePermissionRecords(connector.connector_id, "write");
                                                    self.permissionIdSub = self.permissionId$.subscribe((userpermResp) => {
                                                        self.permissionIdSub.unsubscribe();
                                                        userPermissionId = userpermResp;
                                                        self.a2dAppService.retrieveConnectionDetails(connector, "prioritySync");
                                                        self.a2dAppService.retrieveConnectoionDetailsSub = this.a2dAppService.retrieveConnectoionDetails$.subscribe(
                                                            (response) => {
                                                                self.a2dAppService.retrieveConnectoionDetailsSub.unsubscribe();
                                                                if (response.length > 0) {

                                                                    if (connector.connector_type_value == this.a2dAppService.sharepoint) {
                                                                        connector.authenticated_accesstoken = response[0].ikl_accesstoken;
                                                                        connector.authenticated_refreshtoken = response[0].ikl_refreshtoken;
                                                                        connector.authenticated_conn_detailid = response[0].ikl_connectiondetailid;

                                                                        self.prioritySync(connector, absoluteUrl, fileName, fileUniqueId, userMaxPrivilege.spUserId, userPermissionId, groupCollection, entityConfiguration);

                                                                        self.createSyncStatusRecord(connector, fileActionJsonStr, fileName, fileEntityConfigurationID, syncStatusID);

                                                                    }
                                                                }
                                                                else {
                                                                    this.a2dAppService.connectionDetails = null;
                                                                }

                                                            }, (error) => { }
                                                        );
                                                    })
                                                } else if (self.a2dAppService.isValid(groupNames[0])) {
                                                    self.retrieveGroupDependOnGrpNameCondition(groupNames, connector.connector_id);
                                                    self.sss_GroupCollectionSub = self.sss_GroupCollection$.subscribe((groupCollResp) => {
                                                        self.sss_GroupCollectionSub.unsubscribe();
                                                        groupCollection = groupCollResp;
                                                        self.a2dAppService.retrieveConnectionDetails(connector, "prioritySync");
                                                        this.a2dAppService.retrieveConnectoionDetailsSub = this.a2dAppService.retrieveConnectoionDetails$.subscribe(
                                                            (response) => {
                                                                if (response.length > 0) {

                                                                    if (connector.connector_type_value == this.a2dAppService.sharepoint) {
                                                                        connector.authenticated_accesstoken = response[0].ikl_accesstoken;
                                                                        connector.authenticated_refreshtoken = response[0].ikl_refreshtoken;
                                                                        connector.authenticated_conn_detailid = response[0].ikl_connectiondetailid
                                                                        self.prioritySync(connector, absoluteUrl, fileName, fileUniqueId, userMaxPrivilege.spUserId, userPermissionId, groupCollection, entityConfiguration);
                                                                        self.createSyncStatusRecord(connector, fileActionJsonStr, fileName, fileEntityConfigurationID, syncStatusID);

                                                                    }
                                                                }
                                                                else {
                                                                    this.a2dAppService.connectionDetails = null;
                                                                }

                                                            }, (error) => { }
                                                        );
                                                    })
                                                }
                                                else {
                                                    self.createSyncStatusRecord(connector, fileActionJsonStr, fileName, fileEntityConfigurationID, syncStatusID);
                                                }
                                            }
                                            else {
                                                this.isIkl_FilePrivillagesValid = false;
                                                // if  file entity privillages are "none";
                                                return;

                                            }
                                        })

                                    } else {
                                        this.isIkl_FilePrivillagesValid = false;
                                        // if  file entity privillages are "none";
                                        return;

                                    }

                                })
                            }

                        })

                    }

                });


            }
        } catch (error) {
            throw new Error(`${error.message} - ${functionName}`);
        }
    }


    createSyncStatusRecord(connector: Connector, fileActionJsonStr: string, fileName: string, fileEntityConfigurationID: any, syncStatusID: any) {
        const syncStatusEn = {};
        let self = this;
        try {
            syncStatusEn["ikl_name"] = `Create Action of ${fileName}`;
            syncStatusEn["ikl_actiondetails"] = fileActionJsonStr;
            syncStatusEn["ikl_message"] = 0; // Create
            syncStatusEn["ikl_entityname"] = "sharepointdocumentlocation";
            syncStatusEn["ikl_EntityConfiguration@odata.bind"] = "/ikl_entityconfigurations(" + fileEntityConfigurationID + ")";
            syncStatusEn["ikl_Connector@odata.bind"] = "/ikl_connectors(" + connector.connector_id + ")";
            syncStatusEn["ikl_syncstatus_ikl_sss_syncstatus@odata.bind"] = "/ikl_sss_syncstatuses(" + syncStatusID + ")";


            Xrm.WebApi.createRecord("ikl_sss_syncstatus", syncStatusEn).then((success) => {
                console.log("Sync Status created for the uploaded file");

            }, error => {
                console.log("Error: ", error);
                console.log("Error details: ", error.message);
            })
        }
        catch (err) {
            console.log("Error: ", err);
            console.log("Error details: ", err.message);
        }
    }


    async prioritySync(connector: Connector, relativeUrl: string, fileName: string, fileUniqueId: string, spuserId: string, userPermissionId: string, groupCollection: any, entityConfiguration: EntityConfiguration) {
        let functionName = "prioritySync";
        let self = this;
        try {
            if (self.a2dAppService.isValid(userPermissionId)) {
                switch (connector.auth_type_value) {
                    case self.a2dAppService._sharePointAuthTypes["App"]:
                        let isBreak: any = await self.BreakInheritanceOfSPFolder_App(relativeUrl, fileUniqueId, connector, entityConfiguration);
                        if (isBreak) {
                            relativeUrl = relativeUrl + "/" + fileName;
                            self.AssignGroupOrUserToFolder_App(relativeUrl, fileUniqueId, spuserId, userPermissionId, connector, entityConfiguration);
                        }
                        break;
                    case self.a2dAppService._sharePointAuthTypes["Credential"]:
                        //self.BreakInheritanceOfSPFolder(relativeUrl, connector);
                        break;
                }
            }

            if (self.a2dAppService.isValid(groupCollection)) {
                let groupRecordId: string | undefined = undefined;
                let permissionId: string = '';
                let spgroupId: string = '';

                // Break folder inheritance based on auth type
                switch (connector.auth_type_value) {
                    case self.a2dAppService._sharePointAuthTypes["App"]:
                        let isBreak: any = await self.BreakInheritanceOfSPFolder_App(relativeUrl, fileUniqueId, connector, entityConfiguration);
                        if (isBreak) {
                            //Iterate over groupCollection entities
                            for (let grpOwner = 0; grpOwner < groupCollection.length; grpOwner++) {
                                const entity = groupCollection[grpOwner];
                                spgroupId = entity["ikl_groupid"] || '';
                                groupRecordId = entity["ikl_sssgroupid"] || undefined;
                                permissionId = entity["ikl_permissionid"] || '';

                                switch (connector.auth_type_value) {
                                    case self.a2dAppService._sharePointAuthTypes["App"]:
                                        relativeUrl = relativeUrl + "/" + fileName;
                                        self.AssignGroupOrUserToFolder_App(relativeUrl, fileUniqueId, spgroupId, permissionId, connector, entityConfiguration);
                                        break;
                                    case self.a2dAppService._sharePointAuthTypes["Credential"]:
                                        //self.AssignGroupOrUserToFolder(relativeUrl, spgroupId, permissionId, connector);
                                        break;
                                }
                            }
                        }
                        break;
                    case self.a2dAppService._sharePointAuthTypes["Credential"]:
                        //self.BreakInheritanceOfSPFolder(relativeUrl, connector);
                        break;
                }


            }
        } catch (error) {
            console.log(error.message);
        }
    }

    AssignGroupOrUserToFolder_App(realativeUrl: string, fileUniqueId: string, spgroupId: string, permissionId: string, connector: Connector, entityConfiguration: EntityConfiguration) {
        const functionName: string = "AssignGroupOrUserToFolder";
        let newRelativeUrl: string = '';
        let absoluteURL: string = '';
        let httpOptions: any = {};
        let url: any = "";
        let accessToken: any;
        let self = this;
        try {
            newRelativeUrl = realativeUrl.replace(/'/g, "''");
            newRelativeUrl = newRelativeUrl.replace(connector.absolute_url, "");
            //newRelativeUrl = newRelativeUrl.startsWith("/") ? newRelativeUrl.substring(1) : newRelativeUrl;

            //@ts-ignore
            accessToken = InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.authenticated_accesstoken).DecryptedValue;

            const odataQuery = `_api/web/GetFileById('${fileUniqueId}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${spgroupId},roleDefId=${permissionId})`;
            absoluteURL = connector.absolute_url;
            absoluteURL = absoluteURL.endsWith("/") ? absoluteURL.slice(0, -1) : absoluteURL;
            url = `${absoluteURL}/${odataQuery}`;

            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + accessToken,
                }),
            };

            this.http.post(url, null, httpOptions).subscribe(
                (response) => {

                },
                async (error) => {
                    if (error && error.error && error.error.error_description && error.error.error_description.includes("Invalid JWT token. The token is expired.")) {
                        let tokenResponse: any = await this.generateAccessTokensFromRefreshTokenAuthUser(connector, entityConfiguration);

                        if (tokenResponse && this.a2dAppService.isValid(tokenResponse.access_token)) {
                            //@ts-ignore
                            connector.authenticated_accesstoken = InoEncryption.Encryption.EncryptKey(tokenResponse.access_token);
                            self.AssignGroupOrUserToFolder_App(realativeUrl, fileUniqueId, spgroupId, permissionId, connector, entityConfiguration);
                        }
                    }
                }
            )
        } catch (error: any) {
            console.log(error.message);
        }
    }

    async BreakInheritanceOfSPFolder_App(relativeUrl: string, fileUniqueId: string, connector: Connector, entityConfiguration: EntityConfiguration): Promise<Boolean> {
        // Function level variables
        let newRelativeUrl = '';
        let absoluteURL = connector.absolute_url;
        let httpOptions = {};
        let accessToken: any;
        let self = this;
        try {
            // Replace single quotes and modify the URL
            newRelativeUrl = relativeUrl.replace(/'/g, "''");
            newRelativeUrl = newRelativeUrl.replace(connector.absolute_url, "");
            //newRelativeUrl = newRelativeUrl.startsWith("/") ? newRelativeUrl.substring(1) : newRelativeUrl;

            //@ts-ignore
            accessToken = InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.authenticated_accesstoken).DecryptedValue;

            // OData query to break inheritance
            const odataQuery = `_api/web/GetFileById('${fileUniqueId}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`;

            // Process absolute URL
            absoluteURL = absoluteURL.endsWith("/") ? absoluteURL.slice(0, -1) : connector.absolute_url;

            // Construct full URL
            const url = `${absoluteURL}/${odataQuery}`;
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + accessToken,
                }),
            };

            // Using a promise to handle the observable
            return new Promise<boolean>((resolve) => {
                this.http.post(url, null, httpOptions).subscribe(
                    (response) => {

                        resolve(true);  // Success
                    },
                    async (error) => {
                        if (error && error.error && error.error.error_description && error.error.error_description.includes("Invalid JWT token. The token is expired.")) {
                            let tokenResponse: any = await this.generateAccessTokensFromRefreshTokenAuthUser(connector, entityConfiguration);
                            if (tokenResponse && this.a2dAppService.isValid(tokenResponse.access_token)) {
                                //@ts-ignore
                                connector.authenticated_accesstoken = InoEncryption.Encryption.EncryptKey(tokenResponse.access_token);
                                self.BreakInheritanceOfSPFolder_App(relativeUrl, fileUniqueId, connector, entityConfiguration);
                            }
                        }
                        else {
                            resolve(false);  // Failure
                        }

                    }
                );
            });
        } catch (error) {
            console.log(error.message);
            Promise.resolve(false);
        }
    }


    async getMaxPrivilegeMask(userId: string, fetchCondition: string, connector_id: string) {

        let UserPrvData: UserPrivilegeData = {};
        let maxDepthMask = 0;
        let self = this;
        let queryOptions: any = null;
        let fetchXmlParentYes: string;
        let fetchXmlParentNo: string;
        const combinedRoles = [];
        let fetchXMLTeamUserRootParentYes: string;
        let fetchXMLTeamUserRootParentNo: string;
        try {
            queryOptions = {
                includeFormattedValues: true,
                includeLookupLogicalNames: true,
                includeAssociatedNavigationProperties: true
            };
            fetchXmlParentYes = `
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='role'>
                        <attribute name='name' />
                        <attribute name='businessunitid' />
                        <link-entity name='systemuserroles' from='roleid' to='roleid' visible='false'>
                            <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='User'>
                                <attribute name='domainname' />
                                <filter type='and'>
                                    <condition attribute='accessmode' operator='eq' value='0' />
                                    <condition attribute='azureactivedirectoryobjectid' operator='not-null' />
                                    <condition attribute='isdisabled' operator='eq' value='0' />
                                    <condition attribute='systemuserid' operator='eq' value='${userId}' />
                                </filter>
                                <link-entity name='ikl_connectiondetail' from='ikl_user' to='systemuserid' link-type='inner' alias='SPUser'>
                                    <attribute name='ikl_sharepointuserid' />
                                    <filter type='and'>
                                        <condition attribute='ikl_sharepointuserid' operator='not-null' />
                                        <condition attribute='ikl_connector' operator='eq' value='${connector_id}' />
                                    </filter>
                                </link-entity>
                            </link-entity>
                        </link-entity>
                        <link-entity name='role' from='roleid' to='parentrootroleid' link-type='outer' alias='parentrole'>
                            <link-entity name='roleprivileges' from='roleid' to='roleid' visible='false' alias='roleprv'>
                                <attribute name='privilegedepthmask' />
                                <link-entity name='privilege' from='privilegeid' to='privilegeid' alias='prv'>
                                    <attribute name='name' />
                                    <attribute name='privilegeid' />
                                    <filter type='and'>
                                        ${fetchCondition}
                                    </filter>
                                </link-entity>
                            </link-entity>
                        </link-entity>
                    </entity>
                </fetch>`;

            fetchXmlParentNo = `
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='role'>
                        <attribute name='name' />
                        <attribute name='roleid' />
                        <attribute name='businessunitid' />
                        <link-entity name='systemuserroles' from='roleid' to='roleid' visible='false'>
                            <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='User'>
                                <attribute name='domainname' />
                                <filter type='and'>
                                    <condition attribute='accessmode' operator='eq' value='0' />
                                    <condition attribute='azureactivedirectoryobjectid' operator='not-null' />
                                    <condition attribute='isdisabled' operator='eq' value='0' />
                                    <condition attribute='systemuserid' operator='eq' value='${userId}' />
                                </filter>
                                <link-entity name='ikl_connectiondetail' from='ikl_user' to='systemuserid' link-type='inner' alias='SPUser'>
                                    <attribute name='ikl_sharepointuserid' />
                                    <filter type='and'>
                                        <condition attribute='ikl_sharepointuserid' operator='not-null' />
                                        <condition attribute='ikl_connector' operator='eq' value='${connector_id}' />
                                    </filter>
                                </link-entity>
                            </link-entity>
                        </link-entity>
                        <link-entity name='roleprivileges' from='roleid' to='roleid' visible='false' alias='roleprv'>
                            <attribute name='privilegedepthmask' />
                            <link-entity name='privilege' from='privilegeid' to='privilegeid' alias='prv'>
                                <attribute name='name' />
                                <attribute name='privilegeid' />
                                <filter type='and'>
                                    ${fetchCondition}
                                </filter>
                            </link-entity>
                        </link-entity>
                    </entity>
                </fetch>`;


            fetchXMLTeamUserRootParentYes = `<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >
                <entity name='role' >
                    <attribute name='name' />
                    <attribute name='businessunitid' />
                    <attribute name='roleid' />
                    <link-entity name='teamroles' from='roleid' to='roleid' visible='false' intersect='true'>
                        <link-entity name='team' from='teamid' to='teamid' alias='team'>
                       <attribute name='name' />
                        <attribute name='teamid' /> 
                        <attribute name='businessunitid' />                                              
                        <link-entity name='teammembership' from='teamid' to='teamid' visible='false' intersect='true'>
                        <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='User' >
                        <attribute name='domainname' />
                        <filter type='and' >
                        <condition attribute='azureactivedirectoryobjectid' operator='not-null' />
                        <condition attribute='accessmode' operator='eq' value='0' />
                        <condition attribute='isdisabled' operator='eq' value='0' />   
                        <condition attribute='systemuserid' operator='eq' value='${userId}'/>                             
                        </filter> 
                        <link-entity name='ikl_connectiondetail' from='ikl_user' to='systemuserid' link-type='inner' alias='SPUser'>
                        <attribute name='ikl_sharepointuserid' />
                        <filter type='and'>
                        <condition attribute='ikl_sharepointuserid' operator='not-null' />
                        <condition attribute='ikl_connector' operator='eq' value='${connector_id}' /> 
                        </filter>
                        </link-entity> 
                        <attribute name='fullname' />
                        <attribute name='systemuserid' />
                        <attribute name='domainname' />
                        </link-entity>
                        </link-entity>
                        </link-entity>
                        </link-entity>
                        <link-entity name='role' from='roleid' to='parentrootroleid' link-type='outer' alias='parentrole'>
                        <link-entity name='roleprivileges' from='roleid' to='roleid' visible='false'  alias='roleprv'>
                        <attribute name='privilegedepthmask' />
                        <link-entity name='privilege' from='privilegeid' to='privilegeid' alias='prv' >
                        <attribute name='name' />
                        <attribute name='privilegeid' />
                        <filter type='and' >                                               
                        ${fetchCondition}                                                                                 
                        </filter>
                        </link-entity>
                        </link-entity>
                        </link-entity>
                        </entity>
            </fetch>`;

            fetchXMLTeamUserRootParentNo = `<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >
                <entity name='role'>
                    <attribute name='name' />
                    <attribute name='businessunitid' />
                    <attribute name='roleid' />
                    <link-entity name='teamroles' from='roleid' to='roleid' visible='false' intersect='true'>
                        <link-entity name='team' from='teamid' to='teamid' alias='team'>
                       <attribute name='name' />
                        <attribute name='teamid' /> 
                        <attribute name='businessunitid' />                                               
                        <link-entity name='teammembership' from='teamid' to='teamid' visible='false' intersect='true'>
                        <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='User' >
                        <attribute name='domainname' />
                        <filter type='and' >
                        <condition attribute='accessmode' operator='eq' value='0' />
                        <condition attribute='azureactivedirectoryobjectid' operator='not-null' />
                        <condition attribute='isdisabled' operator='eq' value='0' />
                        <condition attribute='systemuserid' operator='eq' value='${userId}'/>
                        </filter> 
                        <link-entity name='ikl_connectiondetail' from='ikl_user' to='systemuserid' link-type='inner' alias='SPUser'>
                        <attribute name='ikl_sharepointuserid' />
                        <filter type='and'>
                        <condition attribute='ikl_sharepointuserid' operator='not-null' />
                        <condition attribute='ikl_connector' operator='eq' value='${connector_id}' /> 
                        </filter>
                        </link-entity>
                        <attribute name='fullname' />
                        <attribute name='systemuserid' />
                        <attribute name='domainname' />
                        <filter type='and' >
                        <condition attribute='accessmode' operator='eq' value='0' />
                        </filter>
                        </link-entity>
                        </link-entity>
                        </link-entity>
                        </link-entity>
                        <link-entity name='roleprivileges' from='roleid' to='roleid' visible='false' alias='roleprv'>
                        <attribute name='privilegedepthmask' />
                        <link-entity name='privilege' from='privilegeid' to='privilegeid' alias='prv' >
                        <attribute name='name' />
                        <attribute name='privilegeid' />
                        <filter type='and' >                                                
                       ${fetchCondition}       
                        </filter>
                        </link-entity>
                        </link-entity>
                        </entity>
                        </fetch>`;

            fetchXmlParentYes = encodeURIComponent(fetchXmlParentYes);
            fetchXmlParentNo = encodeURIComponent(fetchXmlParentNo);

            // teams
            fetchXMLTeamUserRootParentYes = encodeURIComponent(fetchXMLTeamUserRootParentYes);
            fetchXMLTeamUserRootParentNo = encodeURIComponent(fetchXMLTeamUserRootParentNo);
            // Await the response from the web API call

            let response1 = await self.a2dAppService.webApi.retrieveMultiple("roles", "fetchXml=" + fetchXmlParentYes, queryOptions);
            let securityRolesParentYes = self.a2dAppService.extractResponse(response1);


            let response2 = await self.a2dAppService.webApi.retrieveMultiple("roles", "fetchXml=" + fetchXmlParentNo, queryOptions);
            let securityRolesParentNo = self.a2dAppService.extractResponse(response2);

            //teams
            let responseTeam1 = await self.a2dAppService.webApi.retrieveMultiple("roles", "fetchXml=" + fetchXMLTeamUserRootParentYes, queryOptions);
            let securityRolesTeamParentYes = self.a2dAppService.extractResponse(responseTeam1);


            let responseTeam2 = await self.a2dAppService.webApi.retrieveMultiple("roles", "fetchXml=" + fetchXMLTeamUserRootParentNo, queryOptions);
            let securityRolesTeamParentNo = self.a2dAppService.extractResponse(responseTeam2);


            if (
                self.a2dAppService.isValid(securityRolesParentYes) &&
                self.a2dAppService.isValid(securityRolesParentYes.length) &&
                securityRolesParentYes.length > 0
            ) {
                combinedRoles.push(...securityRolesParentYes);

            }

            if (
                self.a2dAppService.isValid(securityRolesParentNo) &&
                self.a2dAppService.isValid(securityRolesParentNo.length) &&
                securityRolesParentNo.length > 0
            ) {
                combinedRoles.push(...securityRolesParentNo);

            }

            //teams
            if (
                self.a2dAppService.isValid(securityRolesTeamParentYes) &&
                self.a2dAppService.isValid(securityRolesTeamParentYes.length) &&
                securityRolesTeamParentYes.length > 0
            ) {
                combinedRoles.push(...securityRolesTeamParentYes);

            }

            if (
                self.a2dAppService.isValid(securityRolesTeamParentNo) &&
                self.a2dAppService.isValid(securityRolesTeamParentNo.length) &&
                securityRolesTeamParentNo.length > 0
            ) {
                combinedRoles.push(...securityRolesTeamParentNo);

            }

            if (combinedRoles.length > 0) {
                const maxColl = combinedRoles.map(entity => ({
                    max: entity['roleprv.privilegedepthmask'] ?
                        Number(entity['roleprv.privilegedepthmask']) : 0
                })).reduce((prev, curr) => {
                    return { max: Math.max(prev.max, curr.max) };
                }, { max: 0 });

                maxDepthMask = maxColl.max;

                const buId = combinedRoles[0]['_businessunitid_value']
                    ? combinedRoles[0]['_businessunitid_value'].toString()
                    : "";
                const email = combinedRoles[0]['User.domainname']
                    ? combinedRoles[0]['User.domainname'].toString()
                    : "";
                const spUserId = combinedRoles[0]['SPUser.ikl_sharepointuserid']
                    ? combinedRoles[0]['SPUser.ikl_sharepointuserid'].toString()
                    : "";

                UserPrvData.BUId = buId;
                UserPrvData.MaxDepthMask = maxDepthMask;
                UserPrvData.email = email;
                UserPrvData.spUserId = spUserId;
            } else {
                UserPrvData.MaxDepthMask = -1;
            }

            self.userMaxPrev$.next(UserPrvData);
        }
        catch (err) {
            self.userMaxPrev$.next(null);
        }
    }

    getGroupNames(entityName: string, userMaxPrivilege: UserPrivilegeData, connector: Connector, retryCount: any): string[] {
        const functionName = "GetGroupNames";
        let searchByName = "";
        const groupNames: string[] = ["", ""];
        let self = this;


        try {
            switch (userMaxPrivilege.MaxDepthMask) {
                case 1:
                    groupNames[0] = "user";
                    groupNames[1] = "user";
                    break;

                case 2:
                    searchByName = "write_BU";
                    groupNames[0] = `ikl_${entityName}_${searchByName}_${userMaxPrivilege.BUId}`;
                    groupNames[1] = `ikl_${connector.sharepoint_site_id}_${entityName}_${searchByName}_${userMaxPrivilege.BUId}`;
                    break;

                case 4:
                    searchByName = "write_PC_BU";
                    groupNames[0] = `ikl_${entityName}_${searchByName}_${userMaxPrivilege.BUId}`;
                    groupNames[1] = `ikl_${connector.sharepoint_site_id}_${entityName}_${searchByName}_${userMaxPrivilege.BUId}`;
                    break;

                case 8:
                    searchByName = "write";
                    groupNames[0] = `ikl_${entityName}_${searchByName}`;
                    groupNames[1] = `ikl_${connector.sharepoint_site_id}_${entityName}_${searchByName}`;
                    break;
            }


        } catch (error) {
            console.error(`${functionName} : Error Message ${error.message}`);

        }


        return groupNames;
    }


    retrievePermissionRecords(connectorId: string, permissionName: string) {
        // Function level variable declarations
        const functionName = "RetrievePermissionRecords : ";
        let permissionId = "";
        let fetchXML = "";
        let self = this;
        let queryOptions: any = null;
        let extractedResponse: any;
        try {
            queryOptions = {
                includeFormattedValues: true,
                includeLookupLogicalNames: true,
                includeAssociatedNavigationProperties: true
            }

            fetchXML = `
        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='ikl_ssspermission'>
                <attribute name='ikl_name' />
                <attribute name='createdon' />
                <attribute name='ikl_permissionid' />
                <attribute name='ikl_ssspermissionid' />
                <order attribute='ikl_name' descending='false' />
                <filter type='and'>
                    <condition attribute='ikl_name' operator='like' value='%${permissionName}%' />
                </filter>
                <link-entity name='ikl_connector' from='ikl_connectorid' to='ikl_connector' link-type='inner' alias='ad'>
                    <filter type='and'>
                        <condition attribute='ikl_connectorid' operator='eq' value='${connectorId}' />
                    </filter>
                </link-entity>
            </entity>
        </fetch>`;
            fetchXML = encodeURIComponent(fetchXML);
            self.a2dAppService.webApi.retrieveMultiple("ikl_ssspermissions", "fetchXml=" + fetchXML, queryOptions).then(
                (response) => {
                    extractedResponse = self.a2dAppService.extractResponse(response);
                    if (self.a2dAppService.isValid(extractedResponse)) {
                        self.a2dAppService.isValid(extractedResponse[0].ikl_permissionid) ? self.permissionId$.next(extractedResponse[0].ikl_permissionid) : "";
                    } else {
                        self.permissionId$.next("");
                    }
                },
                (error) => {
                    self.permissionId$.next(error);
                }
            )
        } catch (error) {
            console.log(`${functionName} : error message ${error.message}`);
        }
    }

    retrieveGroupDependOnGrpNameCondition(groupNames: string[], connectorId: string) {
        // Function level variables
        const functionName = "RetrieveGroupDependOnGrpNameCondition";
        let fetchXML = "";
        let groupNameCondition = "";
        let self = this;
        let queryOptions: any = null;
        try {
            queryOptions = {
                includeFormattedValues: true,
                includeLookupLogicalNames: true,
                includeAssociatedNavigationProperties: true
            }

            groupNames.forEach((group) => {
                groupNameCondition += `<condition attribute='ikl_name' operator='eq' value='${group}' />`;
            });


            // Generate the FetchXML query
            fetchXML = `
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='ikl_sssgroup'>
                    <attribute name='ikl_name' />
                    <attribute name='ikl_businessunit' />
                    <attribute name='ikl_sssgroupid' />
                    <attribute name='ikl_user' />
                    <attribute name='ikl_groupid' />
                    <attribute name='ikl_permissionid' />
                    <order attribute='ikl_name' descending='false' />
                    <filter type='and'>
                        <filter type='or'>
                            ${groupNameCondition}
                        </filter>
                    </filter>
                    <link-entity name='ikl_entityconfiguration' from='ikl_entityconfigurationid' to='ikl_entityconfiguration' link-type='inner' alias='ad'>
                        <filter type='and'>
                            <condition attribute='ikl_connector' operator='eq' value='${connectorId}' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                    </link-entity>
                </entity>
            </fetch>`;

            fetchXML = encodeURIComponent(fetchXML);

            self.a2dAppService.webApi.retrieveMultiple("ikl_sssgroups", "fetchXml=" + fetchXML, queryOptions).then(
                (response) => {
                    response = self.a2dAppService.extractResponse(response);
                    self.sss_GroupCollection$.next(response);
                },
                (error) => {
                    self.sss_GroupCollection$.next(error);
                }
            )
        } catch (error) {
            console.log(`${functionName} : error message ${error.message}`);
        }
    }


    async retrieveSyncStatus(
        entConfId: string, // Using string to match GUIDs
        connector: any // Use the appropriate type for Connector
    ): Promise<any> { // Replace 'any' with the actual type for EntityCollection
        const functionName = "RetrieveSyncStatus";
        let syncColl: any = null; // Replace 'any' with the actual type for EntityCollection
        let fetchXML = "";
        let queryOptions: any = null;
        let self = this;

        try {
            fetchXML =
                "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
                "                    <entity name='ikl_sss_syncstatus'>" +
                "                        <attribute name='ikl_name' />" +
                "                        <attribute name='ikl_message' />" +
                "                        <attribute name='ikl_entityname' />" +
                "                        <attribute name='ikl_entityconfiguration' />" +
                "                        <attribute name='ikl_actiondetails' />" +
                "                        <attribute name='ikl_sss_syncstatusid' />" +
                "                        <order attribute='ikl_name' descending='false' />" +
                "                        <filter type='and'>" +
                "                            <condition attribute='ikl_syncstatus' operator='null' />" +
                "                            <condition attribute='ikl_name' operator='eq' value='Create Action of ikl_file - " + connector.name + "' />" +
                "                        </filter>" +
                "                        <link-entity name='ikl_entityconfiguration' from='ikl_entityconfigurationid' to='ikl_entityconfiguration' link-type='inner' alias='ad'>" +
                "                            <filter type='and'>" +
                "                                <condition attribute='ikl_entityconfigurationid' operator='eq' value='" + entConfId + "' />" +
                "                            </filter>" +
                "                        </link-entity>" +
                "                    </entity>" +
                "                </fetch>";

            queryOptions = {
                includeFormattedValues: true,
                includeLookupLogicalNames: true,
                includeAssociatedNavigationProperties: true
            };

            fetchXML = encodeURIComponent(fetchXML);
            // Await the response from the web API call


            self.a2dAppService.webApi.retrieveMultiple("ikl_sss_syncstatuses", "fetchXml=" + fetchXML, queryOptions).then(
                (response) => {
                    syncColl = self.a2dAppService.extractResponse(response);
                    if (self.a2dAppService.isValid(syncColl)) {
                        self.syncStatues$.next(syncColl[0].ikl_sss_syncstatusid);
                    }
                    else {
                        self.syncStatues$.next("");
                    }

                }, (error) => {
                    self.syncStatues$.error(error);
                });

        } catch (err) {
            console.error(`${functionName} - Error:`, err);
        }
    }


    async retrieveSharePointLocation(
        regardingReference: any,
        searchByName: boolean,
        parentLocationId: string = ''
    ) {
        const functionName = "retrieveSharePointLocation";
        let sharePointLocation: any | null = null;
        let fetchXML = '';
        let condition = '';
        let self = this;
        let queryOptions: any;
        try {


            // Create the condition
            if (searchByName) {
                condition = `<condition attribute='relativeurl' operator='eq' value='${regardingReference.logicalName}' />
                             <condition attribute='parentsiteorlocation' operator='eq' value='${parentLocationId}' />`;
            } else {
                condition = `<condition attribute='regardingobjectid' operator='eq' value='${regardingReference.id}' />
                             <condition attribute='sitecollectionid' operator='eq' value='${parentLocationId}' />`;
            }

            // Create Fetch
            fetchXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>" +
                "                            <entity name='sharepointdocumentlocation'>" +
                "                                <attribute name='name' />" +
                "                                <attribute name='relativeurl' />" +
                "                                <order attribute='createdon' descending='true' />" +
                "                                <filter type='and'>" +
                "                                    <condition attribute='locationtype' operator='eq' value='0' />" +
                "                                    <condition attribute='servicetype' operator='eq' value='0' />" +
                "                                    " + condition + "" +
                "                                    <condition attribute='statecode' operator='eq' value='0' />" +
                "                                </filter>" +
                "                            </entity>" +
                "                        </fetch>";

            // Retrieve data
            queryOptions = {
                includeFormattedValues: true,
                includeLookupLogicalNames: true,
                includeAssociatedNavigationProperties: true
            }
            fetchXML = encodeURIComponent(fetchXML);

            self.a2dAppService.webApi.retrieveMultiple("sharepointdocumentlocations", "fetchXml=" + fetchXML, queryOptions).then(
                (response: any) => {
                    response = self.a2dAppService.extractResponse(response);
                    let documentLocations = response;
                    if (documentLocations && documentLocations.length > 0) {
                        sharePointLocation = documentLocations[0].sharepointdocumentlocationid;
                        self.spdocloc$.next(sharePointLocation);
                    }
                    else {
                        self.spdocloc$.next(null);
                    }
                },
                (error) => {
                    self.spdocloc$.next(error);
                }
            );


        } catch (error) {
            const errorMessage = error.message || error.toString();
            console.log(error)
        }

        return sharePointLocation;
    }

    async retrieveSecurityMetadata(
        entityRoles: EntityAndRoles[],
        securityPrivilegesRef: SecurityPrivilege[]
    ) {
        const functionName = "RetrieveSecurityMetadata: ";
        let generatedFetch = "";
        let self = this;
        try {
            for (const item of entityRoles) {
                const securityPrivileges: SecurityPrivilege[] = [];

                let entityMetaData = "EntityDefinitions?$select=Privileges,PrimaryNameAttribute,EntitySetName&$filter=LogicalName eq '" + item.EntityName + "'";
                self.a2dAppService.webApi.retrieveMultiple(entityMetaData, null, null).then(
                    (response) => {
                        response = this.a2dAppService.extractResponse(response);

                        let securityPrivilege = response[0].Privileges;


                        for (let index = 0; index < securityPrivilege.length; index++) {
                            if (securityPrivilege[index].PrivilegeType === "Write") { // Adjust to match your PrivilegeType\
                                let privilege: SecurityPrivilege = {};

                                privilege.CanBeBasic = securityPrivilege[index].CanBeBasic;
                                privilege.CanBeDeep = securityPrivilege[index].CanBeDeep;
                                privilege.CanBeEntityReference = securityPrivilege[index].CanBeEntityReference;
                                privilege.CanBeGlobal = securityPrivilege[index].CanBeGlobal;
                                privilege.CanBeLocal = securityPrivilege[index].CanBeLocal;
                                privilege.PrivilegeId = securityPrivilege[index].PrivilegeId;
                                privilege.Name = securityPrivilege[index].Name;
                                privilege.PrivilegeType = securityPrivilege[index].PrivilegeType;

                                securityPrivileges.push(privilege);
                                securityPrivilegesRef.push(privilege);

                            }
                        }
                        item.securityPrivilege = securityPrivileges;
                        generatedFetch = self.generateFetch(item.securityPrivilege); // Implement this function as needed

                        if (self.a2dAppService.isValid(generatedFetch)) {
                            self.entitySecurityMetadataFetch$.next(generatedFetch);
                        }
                        else {
                            self.entitySecurityMetadataFetch$.next("");
                        }

                    },
                    (error) => {
                        self.entitySecurityMetadataFetch$.next(generatedFetch);
                    }
                );
            }
        } catch (error) {
            console.log("err");
        }

    }

    generateFetch(securityPrivilegeMetadatas: SecurityPrivilege[]): string {
        const functionName = "GenerateFetch";
        let fetchXML: string = "";

        try {
            if (securityPrivilegeMetadatas.length > 0) {
                fetchXML = securityPrivilegeMetadatas.map(securityPrivilegeMetadata => {
                    return `<condition attribute='privilegeid' operator='eq' value='${securityPrivilegeMetadata.PrivilegeId}'></condition>`;
                }).join('');
            }
        } catch (error) {
            // Handle specific errors if necessary
            console.log("Err" + error);
        }

        return fetchXML;
    }

    //#endregion

    /**
    * In case of bulk file upload from A2D increase the wait time
    */
    private getAdditionalTimeValue(count: number): number {
        let functionName: string = 'getAdditionalTimeValue';
        let additionalWaitTime: number = 1;
        try {
            if (count <= 50) {
                additionalWaitTime = 1;
            }
            else if (count <= 100) {
                additionalWaitTime = 2;
            }
            else if (count <= 200) {
                additionalWaitTime = 3;
            }
            else if (count <= 300) {
                additionalWaitTime = 4;
            }
            else if (count <= 400) {
                additionalWaitTime = 5;
            }
            else if (count <= 500) {
                additionalWaitTime = 6;
            }
            else {
                additionalWaitTime = 10;
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return additionalWaitTime;
    }

    /**
       * Business Logic for uploading file
       * @param workItems
       * @param uploadPath
       * @param selectedConnectorTab
       * @param selectedEntityConfiguration
       * @param source
       */
    async uploadFileToSPHomeGrid(workItems: any, uploadPath: string, selectedConnectorTab: Connector, selectedEntityConfiguration: any, source: string, runningCount: number, count: any) {
        let functionName: string = "uploadFileToSPHomeGrid";
        try {
            let self = this;
            //Remove the subsite component from the Path
            let subSite: string = this.utilityService.getSharePointSubSite(selectedConnectorTab.absolute_url);
            // for (let value = 0; value < selectedEntityConfiguration.length; value++) {

            //Shreyas 26 May 2022.
            // if (selectedEntityConfiguration.length < 2) {
            //     runningCount = 0;
            //     uploadPath = selectedEntityConfiguration[runningCount].folder_path;
            //     uploadPath = this.utilityService.clearSubSiteFromPath(subSite, uploadPath);
            // }
            // else {

            if (count > runningCount) {
                uploadPath = selectedEntityConfiguration[runningCount].folder_path;
                uploadPath = this.utilityService.clearSubSiteFromPath(subSite, uploadPath);
            }

            // }
            self.a2dAppService.currentCount = 0;

            let fileEntityConfigurationID: any = "";
            if (self.a2dAppService.isSharePointSecuritySyncLicensePresent && selectedConnectorTab.isSyncRecords) {
                fileEntityConfigurationID = await self.getIklFileEntityConfiguration(selectedConnectorTab);
            }

            if (workItems.length > 0 && count > runningCount) {

                for (let i = 0; i < workItems.length + this.a2dAppService.IgnoreFileCount; i++) {
                    //this.modalService.filedisplayMessage = 'File processing ' + " " + (i + 1) + " out of " + workItems.length;
                    if (runningCount == 0 && i == 0) {
                        this.modalService.displayMessage = this.a2dAppService.labelsMultiLanguage['recordprocessing'] + " " + (runningCount + 1) + "/" + selectedEntityConfiguration.length;
                        if (this.a2dAppService.openUploadStatusModal == false)
                            this.modalService.openUploadStatus(this.modalService.displayMessage, (onClose) => {
                            });
                        this.a2dAppService.openUploadStatusModal = true;
                    }
                    if (this.a2dAppService.isValid(workItems[i])) {
                        self.uploadFilesStartSub = new Subscription();
                        let decryptedToken: any = "";
                        let workItem = workItems[i];
                        let file = workItem.file;
                        let path = self.utilityService.formatNameWithOutSlash(workItem.path);
                        let name = self.utilityService.formatFileName(workItem.file.name, selectedConnectorTab.connector_type_value);
                        path = source == "UploadFolder" ? `${uploadPath}${path}` : uploadPath;
                        //reading base64
                        // self.modalService.filedisplayMessage = 'File processing ' + " " + (i + 1) + " out of " + workItems.length;
                        //self.UploadFiles(name, path, base64, selectedConnectorTab, selectedEntityConfiguration);
                        //Shrujan 13 feb 23 Added new method to upload files below 250 Mb size.
                        // self.uploadSPFiles(name, path, base64, selectedConnectorTab, selectedEntityConfiguration, decryptedToken);//shrujan
                        //base64 = "";
                        //Shrujan 09 Aug 23 Added new method to upload files below 1.5 GB size.                               
                        const uploadResponseHome: any = await self.uploadFilesSPmain(file, path, name, selectedConnectorTab, selectedEntityConfiguration);

                        if (self.a2dAppService.isValid(uploadResponseHome)) {
                            self.modalService.fileUploadingPercentage = "";
                            if (uploadResponseHome.status == true || uploadResponseHome.status == "true") {
                                self.a2dAppService.currentCount = self.a2dAppService.currentCount + 1;
                                self.a2dAppService.SuccessFileCount = self.a2dAppService.SuccessFileCount + 1;
                                self.a2dAppService.SuccessFileNames.push(self.createUploadedFileDetailsObject(uploadResponseHome));

                                if (self.a2dAppService.isValid(fileEntityConfigurationID) && this.isIkl_FilePrivillagesValid) {
                                    let fileLogicalName = "ikl_file";
                                    self.createSyncStatusForFile(uploadResponseHome.FileName, uploadResponseHome.FilePath, uploadResponseHome.FileUniqueId, fileLogicalName, fileEntityConfigurationID, selectedConnectorTab, selectedEntityConfiguration)
                                }
                            }
                            else if (uploadResponseHome.status == false || uploadResponseHome.status == "false") {
                                self.a2dAppService.currentCount = self.a2dAppService.currentCount + 1;
                                self.a2dAppService.ErrorFileCount = self.a2dAppService.ErrorFileCount + 1;
                                self.a2dAppService.ErrorFileNames.push(self.createUploadedFileDetailsObject(uploadResponseHome));
                            }
                            // Based on the Running Count and Count, determine the point to turn off the spinner
                            if (self.a2dAppService.selectedEntityRecords.length > 0) {
                                self.modalService.filedisplayMessage = 'File processing ' + " " + self.a2dAppService.currentCount + " out of " + workItems.length;
                                if (self.a2dAppService.currentCount == workItems.length) {
                                    runningCount++;
                                    self.a2dAppService.onGoCount = 0;
                                    self.a2dAppService.currentCount = 0;
                                    self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['recordprocessing'] + " " + (runningCount + 1) + "/" + self.a2dAppService.selectedEntityRecords.length;


                                    self.uploadFileToSPHomeGrid(workItems, uploadPath, selectedConnectorTab, selectedEntityConfiguration, source, runningCount, count);

                                    // if (runningCount == count - 1) {
                                    //     self.spinnerService.hide();
                                    //     self.modalService.filedisplayMessage = "";
                                    //     self.a2dAppService.openUploadStatusModal = false;
                                    //     if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                                    //         self.modalService.UploadStatusModalRef.hide();
                                    //         self.modalService.openDialogWithInputUploadStatus(self.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onClose) => {
                                    //             self.utilityService.gridData = self.utilityService.createCollectionOfFiles(workItems);
                                    //         });
                                    //     }
                                    //     else {
                                    //         self.modalService.UploadStatusModalRef.hide();
                                    //         self.utilityService.gridData = self.utilityService.createCollectionOfFiles(workItems);
                                    //     }
                                    // }
                                }
                            }
                        }
                        else {
                            if (self.a2dAppService.selectedEntityRecords.length > 0) {
                                self.modalService.filedisplayMessage = 'File processing ' + " " + self.a2dAppService.currentCount + " out of " + workItems.length;
                                if (self.a2dAppService.currentCount == workItems.length) {
                                    runningCount++;
                                    self.a2dAppService.onGoCount = 0;
                                    self.a2dAppService.currentCount = 0;
                                    self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['recordprocessing'] + " " + (runningCount + 1) + "/" + self.a2dAppService.selectedEntityRecords.length;
                                    self.uploadFileToSPHomeGrid(workItems, uploadPath, selectedConnectorTab, selectedEntityConfiguration, source, runningCount, count);
                                    // if (runningCount == count - 1) {
                                    //     self.spinnerService.hide();
                                    //     self.modalService.filedisplayMessage = "";
                                    //     self.a2dAppService.openUploadStatusModal = false;
                                    //     if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                                    //         self.modalService.UploadStatusModalRef.hide();
                                    //         self.modalService.openDialogWithInputUploadStatus(self.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onClose) => {
                                    //             self.utilityService.gridData = self.utilityService.createCollectionOfFiles(workItems);
                                    //         });
                                    //     }
                                    //     else {
                                    //         self.modalService.UploadStatusModalRef.hide();
                                    //         self.utilityService.gridData = self.utilityService.createCollectionOfFiles(workItems);
                                    //     }
                                    // }
                                }
                            }
                        }

                    }
                }
            }
            else {
                this.spinnerService.hide();
                if (self.a2dAppService.IgnoreFileCount > 0 || self.a2dAppService.ErrorFileCount > 0) {
                    //self.modalService.UploadStatusModalRef.hide();
                    this.modalService.openDialogWithInputUploadStatus(this.a2dAppService.labelsMultiLanguage['uploadfinalstatus'], (onCloseClick) => {
                        this.utilityService.gridData = self.utilityService.createCollectionOfFiles(workItems);
                    });
                }
                else {
                    self.modalService.UploadStatusModalRef.hide();
                    this.utilityService.gridData = self.utilityService.createCollectionOfFiles(workItems);
                }
            }
        }
        catch (error) {
            //console.log("uploadFileToSPHomeGrid ERROR : " + error);
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Create SharePoint Folders
     * @param folders
     * @param connector
     * @param entityConfiguration
     */
    createFolders(folders: any, connector: Connector, entityConfiguration: EntityConfiguration): void {
        let functionName = "createFolders";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        let rootPath: any = "";
        try {
            //Remove the sub site reference from the path
            let subSite: string = this.utilityService.getSharePointSubSite(connector.absolute_url);
            folders["UploadPath"] = (this.utilityService.clearSubSiteFromPath(subSite, folders["UploadPath"]));
            fileDetail.file_name = folders["Folders"];
            fileDetail.path = folders["UploadPath"];
            entityName = this.a2dAppService.currentEntityName;
            if (this.a2dAppService.selectedEntityRecords.length == 0) {
                recordId = this.a2dAppService.currentEntityId;
            }
            else {
                for (let z = 0; z < this.a2dAppService.selectedEntityRecords.length; z++) {
                    let id: any = this.a2dAppService.selectedEntityRecords[z].replace(/-/g, "");
                    if (entityConfiguration.folder_path.indexOf(id) > -1) {
                        recordId = id;
                    }
                }
            }
            this.createFolders$ = new Subject<any>();
            //Create the Document location and also the folder in CRM
            let object = {
                "MethodName": "createfolder",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "AdditionalDetailsJSON": JSON.stringify(folders),
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration)
            }
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "CreateFolders", entityConfiguration);
            this.createFoldersSub = this.a2dAppService.createFolders$.subscribe(
                (response) => {
                    this.createFoldersSub.unsubscribe();
                    if (response["status"] == "true" || response["status"] == true) {
                        response["entity"] = response["entityConfiguration"];
                        response["isExist"] = "Succeed";
                        this.createFolders$.next(response);
                    }
                    else if (response["status"] == "false" || response["status"] == false) {
                        response["entity"] = response["entityConfiguration"];
                        response["isExist"] = "Failed";
                        this.createFolders$.next(response);
                        if (this.modalService.isOpen == false) {
                            this.modalService.isOpen = true;
                            this.modalService.openErrorDialog(this.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                        this.spinnerService.hide();
                    }
                    if (this.a2dAppService.selectedEntityRecords.length == 0) {
                        if (!folders["UploadPath"].startsWith("/")) {
                            rootPath = "/" + folders["UploadPath"];
                        } else {
                            rootPath = folders["UploadPath"];
                        }

                        //#Added 23/09/2019
                        // On Edge, if trying to update empty folder with no files in it, it doesn't upload but the breadcrumb updates with undefined. So checked condition to avoid that.
                        if (this.a2dAppService.isValid(fileDetail) && this.a2dAppService.isValid(fileDetail.file_name) && this.a2dAppService.isValid(fileDetail.file_name[0]))
                            // Replacing the file name single quote to two single quotes, because in SharePoint, name with single quote is created with 2 single quotes (eg. One's --> One''s)
                            this.getSharePointData(connector, entityConfiguration, rootPath + "/" + fileDetail.file_name[0], self.gridService.selectedView); // .replace(/'/g, "''")
                    }
                },
                (error) => {
                    if (this.modalService.isOpen == false) {
                        this.modalService.isOpen = true;
                        this.modalService.openErrorDialog(this.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    // this.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Upload files to SP
     * @param fileName
     * @param path
     * @param base64
     * @param connector
     * @param entityConfiguration
     */
    UploadFiles(fileName: string, path: string, base64: string, connector: Connector, entityConfiguration: EntityConfiguration): void {
        let functionName: string = "UploadFiles";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        let divideBase64 = [];
        try {
            this.uploadFile$ = new Subject<any>();
            //entityName = this.a2dAppService.currentEntityName;
            //recordId = this.a2dAppService.currentEntityId;
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            // entityName = this.a2dAppService.EntityConfigurationList[runningCount].entity
            // recordId = this.a2dAppService.EntityConfigurationList[runningCount].currentRecordId;
            // //Create the FileDetail Object
            // divideBase64 = this.divideBase64InNParts(base64,10);

            // if(divideBase64 != null && divideBase64.length>0){
            //     let divideBase64Length = divideBase64.length;

            //     if(divideBase64Length > 0){
            //         fileDetail.base64zero = divideBase64[0];
            //     }else{
            //         fileDetail.base64zero = "";
            //     }
            //     if(divideBase64Length > 1){
            //         fileDetail.base64one = divideBase64[1];
            //     }else{
            //         fileDetail.base64one = "";
            //     }
            //     if(divideBase64Length > 2){
            //         fileDetail.base64two = divideBase64[2];
            //     }else{
            //         fileDetail.base64two = "";
            //     }
            //     if(divideBase64Length > 3){
            //         fileDetail.base64three = divideBase64[3];
            //     }else{
            //         fileDetail.base64three = "";
            //     }
            //     if(divideBase64Length > 4){
            //         fileDetail.base64four = divideBase64[4];
            //     }else{
            //         fileDetail.base64four = "";
            //     }
            //       if(divideBase64Length > 5){
            //         fileDetail.base64five = divideBase64[5];
            //     }else{
            //         fileDetail.base64five = "";
            //     }
            //     if(divideBase64Length > 6){
            //         fileDetail.base64six = divideBase64[6];
            //     }else{
            //         fileDetail.base64six = "";
            //     }
            //     if(divideBase64Length > 7){
            //         fileDetail.base64seven = divideBase64[7];
            //     }else{
            //         fileDetail.base64seven = "";
            //     }
            //     if(divideBase64Length > 8){
            //         fileDetail.base64eight = divideBase64[8];
            //     }else{
            //         fileDetail.base64eight = "";
            //     }
            //     if(divideBase64Length > 9){
            //         fileDetail.base64nine = divideBase64[9];
            //     }else{
            //         fileDetail.base64nine = "";
            //     }
            // }

            fileDetail.file_name = fileName;
            fileDetail.path = path;
            fileDetail.base64 = base64;
            //Clear the value in order to avoid any browser hang issue
            base64 = "";
            //Create the Document location and also the folder in CRM
            let object = {
                "MethodName": "uploadfile",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration),
                // "InBase64Zero": fileDetail.base64zero,
                // "InBase64One": fileDetail.base64one,
                // "InBase64Two": fileDetail.base64two,
                // "InBase64Three": fileDetail.base64three,
                // "InBase64Four": fileDetail.base64four,
                // "InBase64Five": fileDetail.base64five,
                // "InBase64Six": fileDetail.base64six,
                // "InBase64Seven": fileDetail.base64seven,
                // "InBase64Eight": fileDetail.base64eight,
                // "InBase64Nine": fileDetail.base64nine,
            }
            //Clear the value in order to avoid any browser hang issue
            fileDetail = {};
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "UploadFile");
            this.uploadFileSub = this.a2dAppService.uploadFile$.subscribe(
                (response) => {
                    this.a2dAppService.currentCount = this.a2dAppService.currentCount + 1;
                    self.uploadFile$.next(response);
                },
                (error) => {
                    self.uploadFile$.next({ "status": false });
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    divideBase64InNParts(base64Text: string, divisionNumber) {
        let functionName: string = "divideBase64InNParts";
        try {
            const len = base64Text.length / divisionNumber;
            const creds = base64Text.split("").reduce((acc, val) => {
                let { res, currInd } = acc;
                if (!res[currInd] || res[currInd].length < len) {
                    res[currInd] = (res[currInd] || "") + val;
                } else {
                    res[++currInd] = val;
                };
                return { res, currInd };
            }, {
                res: [],
                currInd: 0
            });
            return creds.res;
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    // ShrujanB
    /**Shrujan 13 feb 22 Added new method to upload files below 250 Mb size.
     * This methos
     * @param name 
     * @param selectedConnectorTab 
     * @param selectedEntityConfiguration 
     * @param fileObj 
     */
    uploadSPFiles(name: any, path: any, base64: any, selectedConnectorTab: any, selectedEntityConfiguration: any, decryptedToken: any) {
        let functionName: string = "uploadSPFiles";
        let sharePointSite: any;
        let folderPath: string;
        let encryptedToken: any;
        let httpOptions = {};
        let response = {};
        let requestUrl: any;
        let self = this;
        let isOverride: boolean = false;
        let stringIsOverride: string;
        let folderServerRelativeURL: string;
        let firstCharOfPath: string;
        let fileName: string;
        let folderRelativePath: string = "";
        let runningCount = 1;
        let contents: any;

        try {
            if (this.a2dAppService.isValid(selectedConnectorTab) && this.a2dAppService.isValid(selectedEntityConfiguration)) {
                isOverride = selectedEntityConfiguration.isOverride;
                folderPath = selectedEntityConfiguration.folder_path;
                stringIsOverride = isOverride ? "true" : "false";
                this.uploadFile$ = new Subject<any>();
                fileName = name;
                folderRelativePath = path.replace(/'/g, "''");
                fileName = fileName.replace(/'/g, "''");
                firstCharOfPath = path.charAt(0);

                //Checking if path have "/"" character at first position if contains then remove because it is not supports inrequest url (GetFolderByServerRelativeUrl).
                if (firstCharOfPath == "/") {
                    folderServerRelativeURL = encodeURIComponent(folderRelativePath.substring(1, folderRelativePath.length));
                }
                else {
                    folderServerRelativeURL = encodeURIComponent(folderRelativePath);
                }

                sharePointSite = selectedConnectorTab.absolute_url;
                requestUrl = sharePointSite + "/_api/Web/GetFolderByServerRelativeUrl('" + folderServerRelativeURL + "')/Files/add(overwrite='" + stringIsOverride + "', url='" + fileName + "')?$expand=ListItemAllFields";

                contents = this.base64ToArrayBuffer(base64);
                contents.byteLength;

                httpOptions = {
                    headers: new HttpHeaders({
                        "accept": "application/json;odata=verbose",
                        "Authorization": "Bearer " + decryptedToken,
                        "content-length": contents.byteLength,
                        'Content-Type': 'multipart/form-data'
                    }),
                };
                this.http.post(requestUrl, contents, httpOptions).subscribe
                    (result => {
                        response["FileName"] = name;
                        response["FilePath"] = folderPath;
                        response["status"] = "true";
                        this.a2dAppService.currentCount = this.a2dAppService.currentCount + 1;
                        self.uploadFile$.next(response);
                    },
                        error => {
                            if ('error' in error.error) {
                                // if flie alredy exist  error reuturn  2130575257 code
                                if (this.a2dAppService.isValid(error.error.error.code) && error.error.error.code.includes("2130575257")) {
                                    let newFilename = name;
                                    let timestamp = new Date().getTime();
                                    let ext = name.split('.').pop(); // get file extension
                                    newFilename = `${name.replace(`.${ext}`, '')}_${timestamp}.${ext}`;
                                    newFilename = newFilename.replace(/'/g, "''");
                                    requestUrl = sharePointSite + "/_api/Web/GetFolderByServerRelativeUrl('" + folderServerRelativeURL + "')/Files/add(overwrite='" + stringIsOverride + "', url='" + newFilename + "')?$expand=ListItemAllFields";
                                    this.http.post(requestUrl, contents, httpOptions).subscribe
                                        (result => {
                                            response["FileName"] = name;
                                            response["FilePath"] = folderPath;
                                            response["status"] = "true";
                                            this.a2dAppService.currentCount = this.a2dAppService.currentCount + 1;
                                            self.uploadFile$.next(response);
                                        },
                                            error => {
                                                //console.error('There was an error!', error);
                                                response["FileName"] = name;
                                                response["FilePath"] = folderPath;
                                                response["status"] = "false";
                                                self.uploadFile$.next(response);
                                            })
                                }
                                else {
                                    response["FileName"] = name;
                                    response["FilePath"] = folderPath;
                                    response["status"] = "false";
                                    self.uploadFile$.next(response);
                                }
                            }
                            else if (error.error.error_description = "Invalid JWT token. The token is expired.") {
                                this.generateAccessTokensFromRefreshToken(selectedConnectorTab, selectedEntityConfiguration);
                                this.generateSPAccessTokenSub = this.generateSPAccessToken$.subscribe(
                                    (response: any) => {
                                        if (response != null && this.a2dAppService.isValid(response.access_token)) {
                                            decryptedToken = response.access_token;
                                            //@ts-ignore
                                            selectedConnectorTab.access_token = InoEncryption.Encryption.EncryptKey(response.access_token);
                                            let newfilename = fileName;
                                            newfilename = newfilename.replace(/'/g, "''");
                                            requestUrl = sharePointSite + "/_api/Web/GetFolderByServerRelativeUrl('" + folderServerRelativeURL + "')/Files/add(overwrite='" + stringIsOverride + "', url='" + newfilename + "')?$expand=ListItemAllFields";
                                            httpOptions = {
                                                headers: new HttpHeaders({
                                                    "accept": "application/json;odata=verbose",
                                                    "Authorization": "Bearer " + response.access_token,
                                                    "content-length": contents.byteLength,
                                                    'Content-Type': 'multipart/form-data'
                                                }),
                                            };
                                            this.http.post(requestUrl, contents, httpOptions).subscribe
                                                (result => {
                                                    response["FileName"] = newfilename;
                                                    response["FilePath"] = folderPath;
                                                    response["status"] = "true";
                                                    this.a2dAppService.currentCount = this.a2dAppService.currentCount + 1;
                                                    self.uploadFile$.next(response);
                                                },
                                                    error => {
                                                        if ('error' in error.error) {
                                                            // if flie alredy exist  error reuturn  2130575257 code
                                                            if (this.a2dAppService.isValid(error.error.error.code) && error.error.error.code.includes("2130575257")) {
                                                                let newFilename = name;
                                                                let timestamp = new Date().getTime();
                                                                let ext = name.split('.').pop(); // get file extension
                                                                newFilename = `${name.replace(`.${ext}`, '')}_${timestamp}.${ext}`;
                                                                newFilename = newFilename.replace(/'/g, "''");
                                                                requestUrl = sharePointSite + "/_api/Web/GetFolderByServerRelativeUrl('" + folderServerRelativeURL + "')/Files/add(overwrite='" + stringIsOverride + "', url='" + newFilename + "')?$expand=ListItemAllFields";
                                                                this.http.post(requestUrl, contents, httpOptions).subscribe
                                                                    (result => {
                                                                        response["FileName"] = name;
                                                                        response["FilePath"] = folderPath;
                                                                        response["status"] = "true";
                                                                        this.a2dAppService.currentCount = this.a2dAppService.currentCount + 1;
                                                                        self.uploadFile$.next(response);
                                                                    },
                                                                        error => {
                                                                            //console.error('There was an error!', error);
                                                                            response["FileName"] = name;
                                                                            response["FilePath"] = folderPath;
                                                                            response["status"] = "false";
                                                                            self.uploadFile$.next(response);
                                                                        })
                                                            }
                                                        }
                                                        else {
                                                            response["FileName"] = fileName;
                                                            response["FilePath"] = folderPath;
                                                            response["status"] = "false";
                                                            self.uploadFile$.next(response);
                                                        }
                                                    })
                                        }
                                        else {
                                            response["FileName"] = fileName;
                                            response["FilePath"] = folderPath;
                                            response["status"] = "false";
                                            self.uploadFile$.next(response);
                                        }
                                    },
                                    (error) => {
                                        response["FileName"] = fileName;
                                        response["FilePath"] = folderPath;
                                        response["status"] = "false";
                                        self.uploadFile$.next(response);
                                    }
                                );
                            }
                            else {
                                response["FileName"] = fileName;
                                response["FilePath"] = folderPath;
                                response["status"] = "false";
                                self.uploadFile$.next(response);
                            }
                        }
                    );
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
            this.a2dAppService.logError('', error.description || error.message, selectedEntityConfiguration, '', null, null);
        }
    }

    base64ToArrayBuffer(base64: string): any {
        var functionName = "base64ToArrayBuffer";
        try {
            var binaryString = window.atob(base64);
            var len = binaryString.length;
            var content = new Uint8Array(len);
            for (var i = 0; i < len; i++) {
                content[i] = binaryString.charCodeAt(i);
            }
            return content.buffer;
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    generateAccessTokensFromRefreshToken(connector: Connector, selectedEntityConfiguration: EntityConfiguration) {
        return new Promise((resolve, reject) => {
            let functionName: string = "generateAccessTokensFromRefreshToken";
            let parameters: any = null;
            let entityName: string = "";
            let recordId: string = "";
            let self = this;
            //let connectorDetail: Connector = {};
            let divideBase64 = [];
            let additionalDetails: {} = {};
            let additionalDetailsJSON: string = "";
            try {
                this.generateSPAccessToken$ = new Subject<any>();
                entityName = this.a2dAppService.currentEntityName;
                recordId = this.a2dAppService.currentEntityId;
                additionalDetails = {
                    "tenant_id": this.a2dAppService.tenantId,
                };
                additionalDetailsJSON = JSON.stringify(additionalDetails);
                let object = {
                    "MethodName": "getaccesstokenfromrefreshtoken",
                    "ConnectorJSON": JSON.stringify(connector),
                    "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                    "EntityName": entityName,
                    "RecordId": recordId,
                    "AdditionalDetailsJSON": additionalDetailsJSON,
                    "EntityConfigurationJSON": JSON.stringify(selectedEntityConfiguration),
                }
                this.a2dAppService.callSharePointCoreAction(object, this.actionName, "GetAceessTokens");
                this.generateSPAccessTokenSub = this.a2dAppService.generateSPAccessToken$.subscribe(
                    (response) => {
                        //self.generateSPAccessToken$.next(response);
                        resolve(response);
                    },
                    (error) => {
                        self.generateSPAccessToken$.next({ "status": false });
                        self.utilityService.throwError(error, functionName);
                    }
                );
            }
            catch (error) {
                this.utilityService.throwError(error, functionName);
            }
        })
    }


    generateAccessTokensFromRefreshTokenAuthUser(connector: Connector, selectedEntityConfiguration: EntityConfiguration) {
        return new Promise((resolve, reject) => {
            let functionName: string = "generateAccessTokensFromRefreshTokenAuthUser";
            let parameters: any = null;
            let entityName: string = "";
            let recordId: string = "";
            let self = this;
            //let connectorDetail: Connector = {};
            let divideBase64 = [];
            let additionalDetails: {} = {};
            let additionalDetailsJSON: string = "";
            try {
                this.generateSPAccessToken$ = new Subject<any>();
                entityName = this.a2dAppService.currentEntityName;
                recordId = this.a2dAppService.currentEntityId;
                additionalDetails = {
                    "tenant_id": this.a2dAppService.tenantId,
                };
                additionalDetailsJSON = JSON.stringify(additionalDetails);
                let object = {
                    "MethodName": "getaccesstokenfromrefreshtokenauthuser",
                    "ConnectorJSON": JSON.stringify(connector),
                    "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                    "EntityName": entityName,
                    "RecordId": recordId,
                    "AdditionalDetailsJSON": additionalDetailsJSON,
                    "EntityConfigurationJSON": JSON.stringify(selectedEntityConfiguration),
                }
                this.a2dAppService.callSharePointCoreAction(object, this.actionName, "GetAceessTokens");
                this.generateSPAccessTokenSub = this.a2dAppService.generateSPAccessToken$.subscribe(
                    (response) => {
                        //self.generateSPAccessToken$.next(response);
                        resolve(response);
                    },
                    (error) => {
                        self.generateSPAccessToken$.next({ "status": false });
                        self.utilityService.throwError(error, functionName);
                    }
                );
            }
            catch (error) {
                this.utilityService.throwError(error, functionName);
            }
        })
    }

    //Shrujan 13 feb 23 for D&D 
    //move files to sharepoint 
    /**
     * This function move draged file to sharepoint folder
     * @param draggedFile 
     * @param destinationfolder 
     * @param connector 
     * @param entityConfiguration 
     */
    MoveFilesToSp(draggedFile: any, destinationfolder: any, connector: Connector, entityConfiguration: EntityConfiguration): void {
        let functionName: string = "MoveFiles";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        let divideBase64 = [];
        try {
            this.moveFile$ = new Subject<any>();
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            fileDetail.file_name = draggedFile;
            fileDetail.path = destinationfolder;
            // let object = {
            //     "MethodName": "movefile",
            //     "ConnectorJSON": JSON.stringify(connector),
            //     "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
            //     "EntityName": entityName,
            //     "RecordId": recordId,
            //     "FileDetailsJSON": JSON.stringify(fileDetail),
            //     "EntityConfigurationJSON": JSON.stringify(entityConfiguration),
            // }
            // this.a2dAppService.callSharePointCoreAction(object, this.actionName, "MoveFile");
            this.moveFileForSharepoint(draggedFile, destinationfolder, connector, entityConfiguration);
            this.moveFileSPSub = this.a2dAppService.moveFile$.subscribe(
                (response) => {
                    self.moveFile$.next(response);
                },
                (error) => {
                    self.moveFile$.next({ "status": false });
                    self.utilityService.throwError(error, functionName);
                }
            );
            //fileDetail = {};
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }


    moveFileForSharepoint(draggedFile: any, destinationfolder: any, connector: Connector, entityConfiguration: EntityConfiguration) {
        let functionName = 'moveFileForSharepoint ()=> '
        let requestUrl: string;
        let decryptedToken: string;
        let fileName: any;
        let drgFile = draggedFile;
        let destFol = destinationfolder;
        try {
            let destinationFol = destinationfolder.replaceAll(/'/g, "''");
            destinationFol = encodeURIComponent(destinationFol);
            destinationFol = destinationFol.replaceAll("+", "%20");
            fileName = encodeURIComponent(draggedFile.FileLeafRef);
            fileName = fileName.replaceAll("+", "%20");
            fileName = fileName.replaceAll(/'/g, "''");

            //@ts-ignore
            decryptedToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;
            requestUrl = `${connector.absolute_url}/_api/web/getFileById('${draggedFile.UniqueId}')/moveto(newurl='${destinationFol + '/' + fileName}')`;
            const headers = new HttpHeaders({
                'Content-Type': 'application/octet-stream',
                "Authorization": "Bearer " + decryptedToken
            });

            this.http.post(requestUrl, null, { headers }).subscribe((response) => {
                console.log('File moved successfully');


            }, async (error: any) => {
                if (error && error.error && error.error['odata.error'] && error.error['odata.error'].code && error.error['odata.error'].code.includes("-2130575257")) {
                    this.moveFile(draggedFile, destinationfolder, connector, entityConfiguration, decryptedToken);
                }

                if (error && error.error && error.error.error_description && error.error.error_description.includes("Invalid JWT token. The token is expired.")) {
                    let tokenResponse: any = await this.generateAccessTokensFromRefreshToken(connector, entityConfiguration);
                    this.acceessToken = tokenResponse ? tokenResponse.access_token : null;

                    if (tokenResponse && this.a2dAppService.isValid(tokenResponse.access_token)) {
                        this.moveFileForSharepoint(drgFile, destFol, connector, entityConfiguration);
                    }
                }
                else {
                    console.error('Error moving file:', error);
                }
            });
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    moveFile(draggedFile: any, destinationfolder: any, connector: Connector, entityConfiguration: EntityConfiguration, decryptedToken: string) {
        let functionName = 'moveFile () => ';
        let arrayBuffer: any;
        let entityConfig: any = [entityConfiguration];
        let fileUniqueId: string = draggedFile.UniqueId;
        let sitePath: any;
        try {
            sitePath = connector.absolute_url.split("/");
            if (sitePath.length > 3) {
                sitePath = '/' + sitePath.splice(3).join('/');
                draggedFile.path_display = draggedFile.path_display.replace(sitePath, '');
                draggedFile.path_display = draggedFile.path_display.startsWith('/') ? draggedFile.path_display.substring(1) : draggedFile.path_display;
                destinationfolder = destinationfolder.replace(sitePath, '');
                destinationfolder = destinationfolder.startsWith('/') ? destinationfolder.substring(1) : destinationfolder;

            }
            let relativeUrl = draggedFile.path_display.trimStart('/');
            let lastSlashIndex = relativeUrl.lastIndexOf('/');

            // Extract the path without the file name
            let newPath = relativeUrl.substring(0, lastSlashIndex);
            this.download$ = new Subject<any>();
            //this.a2dAppService.deleteFileSP$ = new Subject();
            this.downloadSPAngular(newPath, draggedFile.FileLeafRef, connector, entityConfiguration, true, fileUniqueId);
            this.downloadSPSub = this.download$.subscribe(async (downloadResponse) => {
                arrayBuffer = downloadResponse.arrayBuffer;
                if (arrayBuffer != null) {
                    let blob = new Blob([new Uint8Array(arrayBuffer)], { type: 'application/octet-stream' });
                    let file = new File([blob], downloadResponse.fileName);
                    this.uploadFilesSPmain(file, destinationfolder, downloadResponse.fileName, connector, entityConfiguration).then((response: any) => {
                        if (this.a2dAppService.isValid(response)) {
                            if (response.status == true || response.status == "true") {
                                let path: string;
                                if (connector.absolute_url.split("/").length > 3) {
                                    path = sitePath + newPath + '/' + downloadResponse.fileName;
                                }
                                else {
                                    path = newPath + '/' + downloadResponse.fileName;
                                }
                                entityConfig = [entityConfiguration];
                                this.deleteFile(path, downloadResponse.fileName, '', connector, entityConfig, fileUniqueId, true);
                                this.deleteFileSub = this.a2dAppService.deleteFileSP$.subscribe(
                                    (response) => {
                                        if (response["status"] == "true" || response["status"] == true) {
                                            this.moveFile$.next(true);
                                        }
                                    }
                                )
                            }
                        }
                    })
                } error => {
                    this.utilityService.throwError(error, functionName);
                }
            })
        } catch (error) {
            this.utilityService.throwError(error, functionName);
            return false;
        }
    }


    /**
     * Crete note and open the docuign ui - I this function will call the sharepoint core action to get the base64 of
     * @param data
     * @param connector
     * @param entityConfiguration
     * @param buttonName
     */
    createNotesAndOpenUI(data: any, connector: Connector, entityConfiguration: EntityConfiguration, buttonName: string): string {
        let functionName: string = "createNotesAndOpenUI";
        let self = this;
        let count: number = 1;
        let runningCount: number = 1;
        let fileName: string = "";
        let uploadPath: string = "";
        let relativePath: string = "";
        let selectedColl: any = [];
        let blocked: boolean = false;
        let fileExtension: any = "";
        let countOfIgnore: number = 0;
        let base64String = "";
        let currentTimeStamp: Date;
        let splitFileName = [];
        try {
            // This condition is valid when the user clicks on Download button after selecting the record
            if (data != null && data["selectedGridData"].length > 0) {
                count = data["selectedGridData"].length;
                for (let index = 0; index < data["selectedGridData"].length; index++) {
                    relativePath = data["selectedGridData"][index].path_display;
                    fileName = relativePath.split("/").pop();
                    if (data["selectedGridData"][index]["fileType"] == "file") {
                        this.spinnerService.show();
                        // Get the SharePoint Sub Site if any
                        let subSite: string = this.utilityService.getSharePointSubSite(connector.absolute_url);
                        //get the cleared path
                        uploadPath = this.utilityService.clearSubSiteFromPath(subSite, relativePath);
                        //Replace file name from the upload path
                        uploadPath = uploadPath.replace(fileName, "");
                        this.download(uploadPath, fileName, connector, entityConfiguration);
                        this.downloadSPSub = this.download$.subscribe(
                            (response) => {
                                base64String = response["base64"];
                                currentTimeStamp = new Date(this.utilityService.getDateTimeInUserTZ(this.a2dAppService.crmUserTimeZoneParameter["TimeZoneBias"], new Date(), "", ""));
                                splitFileName = response.fileName.split('.');
                                fileName = connector.signed_document_naming == "966620000" ? response.fileName : splitFileName[0] + String(currentTimeStamp.getTime()) + "." + splitFileName[1];
                                this.createNote(entityConfiguration, buttonName, data, base64String, fileName)
                            },
                            (error) => {
                                if (self.modalService.isOpen == false) {
                                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                }
                                self.utilityService.throwError(error, functionName);
                            }
                        );
                    }
                    else {
                        runningCount++;
                        self.spinnerService.hide();
                    }
                }
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return base64String;
    }

    /**
     * Will create a not with the selected document as attachment
     * @param entityConfiguration
     * @param buttonName
     * @param data
     * @param base64String
     */
    createNote(entityConfiguration: any, buttonName: string, data: any, base64String: string, fileName: string) {
        let functionName: string = "createNote";
        let entity = {};
        let actionName = "";
        let req: any = {};
        let dynamicsLookup = "objectid_" + this.a2dAppService.currentEntityName + "@odata.bind";
        let regarding = `/${this.a2dAppService.currentEntityPluralName}(${this.a2dAppService.currentEntityId})`;
        try {
            actionName = (buttonName == "getSignature") ? "ikl_InogicGetSignature" : "ikl_InogicSelfSign";

            entity["subject"] = "Ignore#Ignore Inogic Docusign Integration Note";
            entity["notetext"] = "Ignore#Ignore Inogic Docusign Integration Note";
            entity["filename"] = fileName;
            entity["isdocument"] = true;
            entity["documentbody"] = base64String;
            entity[dynamicsLookup] = regarding

            this.a2dAppService.webApi.create("annotations", entity).then((response: any) => {
                response = this.a2dAppService.extractResponse(response);
                (buttonName == "getSignature") ? this.createGetSignatureReq(actionName, response.id, entityConfiguration, fileName) : this.getLoginUserDetails(actionName, response.id, entityConfiguration, fileName);

            }, (error: any) => {
                this.a2dAppService.logError('', error.description || error.message, entityConfiguration[0], '', null, null);
                this.modalService.isOpen = true;
                this.modalService.openErrorDialog(error.description || error.message, (onOKClick) => { });
                this.utilityService.throwError(error, functionName);
            })
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Call the docusign integration actions to get the envelope url according to the
     * user choice and the action is completed delete the note which we have created earlier
     * @param actionName
     * @param annotationID
     * @param entityConfiguration
     * @param req
     */
    callInoSignAction(actionName: string, annotationID: string, entityConfiguration: any, req: {}) {
        let functionName = "callInoSignAction";
        try {
            this.a2dAppService.webApi.unboundAction(actionName, req).then((response: any) => {
                response = this.a2dAppService.extractResponse(response);
                this.openURLUsingNavigateTo(response.EnvelopeURL);
                this.spinnerService.hide();
                //this.appService._Xrm.openUrl(response.EnvelopeURL, null);
                this.deleteAnnoation(annotationID);

            }, (error: any) => {
                this.deleteAnnoation(annotationID);
                this.a2dAppService.logError('', error.description || error.message, entityConfiguration[0], '', null, null);
                this.modalService.isOpen = true;
                this.modalService.openErrorDialog(error.description || error.message, (onOKClick) => { });
                this.utilityService.throwError(error, functionName);
            });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Open the url
     * @param url
     */
    openEnvelopeURL(url: string) {
        let functionName = "openEnvelopeURL";
        let size;
        let height;
        let windowWidth;
        let windowHeight;
        let webResourceName = "ikl_/Attach2Dynamics/DocuSign.html";
        let crmURL = "";
        let data = "";
        let windowOptions = {}
        try {
            size = (parent.window.innerWidth > 0) ? parent.window.innerWidth : screen.width;
            height = (parent.window.innerHeight > 0) ? parent.window.innerHeight : screen.height;

            //set window width
            windowWidth = "640"; // 610
            //set window height
            windowHeight = "535";

            if (parseFloat(size) < 500) {
                windowWidth = size;
            }

            data = encodeURIComponent(`url=${url}`);
            windowOptions = { height: windowHeight, width: windowWidth };
            //@ts-ignore
            Xrm.Navigation.openWebResource(webResourceName, windowOptions, data)
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * In UCI ope the envelope url using navigate to
     * @param url
     */
    openURLUsingNavigateTo(envelopeUrl: string) {
        let functionName = "openURLUsingNavigateTo";
        let self = this;
        let pageInput = {};
        let navigationOptions = {};
        let webresourceName = "ikl_/Attach2Dynamics/DocuSign.html";
        try {
            pageInput = {
                pageType: "webresource",
                webresourceName: webresourceName,
                data: "url=" + envelopeUrl

            };
            navigationOptions = {
                target: 2,
                height: { value: 80, unit: "%" },
                width: { value: 70, unit: "%" },
                position: 1
            };

            //@ts-ignore
            Xrm.Navigation.navigateTo(pageInput, navigationOptions).then((success: any) => {
                //this.showA2DUI();
            }).catch((error: any) => {
                this.openEnvelopeURL(envelopeUrl);
            });

        }
        catch (error) {
            this.openEnvelopeURL(envelopeUrl);
            // this.utilityService.throwError(error, functionName);
        }
    }

    //showA2DUI() {
    //    let functionName = "showA2DUI";
    //    let a2DDialog = parent.document.getElementById("AlertA2DJs-dialog");
    //    let a2DBackground = parent.document.getElementById("AlertA2D-js-background");
    //    try {
    //        if (a2DDialog != null || a2DBackground != null) {
    //            parent.document.getElementById("AlertA2DJs-dialog").style.zIndex = "1006";
    //            parent.document.getElementById("AlertA2D-js-background").style.zIndex = "1005";
    //        }
    //    }
    //    catch (error) {
    //        this.utilityService.throwError(error, functionName);
    //    }
    //}

    /**
     * Create req with necessary paramter for GetSignature action
     * @param url
     */
    createGetSignatureReq(actionName: string, annotationID: string, entityConfiguration: any, fileName: string) {
        let functionName = "createGetSignatureReq";
        let req: any = {};
        try {
            //Set input parameter
            req.EntityName = this.a2dAppService.currentEntityName;
            req.EntityId = this.a2dAppService.currentEntityId;
            req.FileName = fileName;
            //Get meta data function
            req.getMetadata = function () {
                return {
                    boundParameter: null,
                    parameterTypes: {
                        "EntityName": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                        "EntityId": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                        "FileName": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                    },
                    operationType: 0, //Action
                    operationName: actionName
                };
            };

            this.callInoSignAction(actionName, annotationID, entityConfiguration, req);
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Create Req with necessary parater for self sign
     * @param actionName
     * @param annotationID
     * @param entityConfiguration
     * @param response
     */
    createSelfSign(actionName: string, annotationID: string, entityConfiguration: any, response: any, fileName: string) {
        let functionName = "createSelfSign";
        let req: any = {};
        try {
            //Set input parameter
            req.EntityName = this.a2dAppService.currentEntityName;
            req.EntityId = this.a2dAppService.currentEntityId;
            req.RecipientName = response.fullname;
            req.RecipientEmailId = response.internalemailaddress;
            req.FileName = fileName;
            //Get meta data function
            req.getMetadata = function () {
                return {
                    boundParameter: null,
                    parameterTypes: {
                        "EntityName": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                        "EntityId": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                        "RecipientName": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                        "RecipientEmailId": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                        "FileName": {
                            "typeName": "Edm.String",
                            "structuralProperty": 1 // PrimitiveType
                        },
                    },
                    operationType: 0, //Action
                    operationName: actionName
                };
            };

            this.callInoSignAction(actionName, annotationID, entityConfiguration, req);
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Get login user name and email address this need as paramter for self sign
     * @param actionName
     * @param annotationID
     * @param entityConfiguration
     */
    getLoginUserDetails(actionName: string, annotationID: string, entityConfiguration: any, fileName: string) {
        let functionName: string = "";
        let fetchXml: string = "";
        let queryOptions = {}
        let userId: any;
        try {
            userId = this.a2dAppService._Xrm.getUserId().substring(1, this.a2dAppService._Xrm.getUserId().length - 1);

            this.a2dAppService.webApi.retrieve("systemusers", userId, "?$select=fullname,internalemailaddress").then((response: any) => {
                response = this.a2dAppService.extractResponse(response);
                this.createSelfSign(actionName, annotationID, entityConfiguration, response, fileName);
            },
                (error) => {
                    this.utilityService.throwError(error, functionName);
                });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Delete annoatation
     * @param annotationID
     */
    deleteAnnoation(annotationID: any) {
        let functionName = "deleteAnnoation";
        let recordId = this.a2dAppService.currentEntityId;
        let entityName = this.a2dAppService.currentEntityName;
        try {
            this.a2dAppService.webApi.delete("annotations", annotationID).then((response: any) => {

            }, (error: any) => {
                this.utilityService.throwError(error, functionName);
            });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    /**
    * Download File from SP
    * @param data
    * @param uploadPath
    * @param connector
    * @param entityConfiguration
    */
    downloadSP(data: any, connector: Connector, entityConfiguration: EntityConfiguration, flag: boolean, path?: string): void {
        let functionName: string = "downloadSP";
        let self = this;
        let count: number = 1;
        let runningCount: number = 1;
        let fileName: string = "";
        let uploadPath: string = "";
        let relativePath: string = "";
        let selectedColl: any = [];
        let blocked: boolean = false;
        let fileExtension: any = "";
        let countOfIgnore: number = 0;
        let isOpenDownloadModel: boolean = true;
        let downloadContentResponse: any = {}
        let myCount: number = 1;
        let fileUniqueId: string;
        try {
            // This condition is valid when the user clicks on Download button after selecting the record
            if (data != null && data["selectedGridData"].length > 0) {
                count = data["selectedGridData"].length;
                self.modalService.displayMessage = "";
                self.modalService.fileUploadingPercentage = "";
                //self.modalService.displayMessage = "Downloading files " + runningCount + "/" + count;
                self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['downloadingfile'] + " " + runningCount + "/" + count;
                if (+data["selectedGridData"][0]["File_x0020_Size"] < 200 && data["selectedGridData"].length == 1) {
                    isOpenDownloadModel = false;
                };
                if (flag == true && isOpenDownloadModel == true) {
                    self.modalService.openUploadStatus(self.modalService.displayMessage, (onClose) => {
                    });
                }
                for (let index = 0; index < data["selectedGridData"].length; index++) {
                    relativePath = data["selectedGridData"][index].path_display;
                    fileName = relativePath.split("/").pop();
                    fileUniqueId = data["selectedGridData"][index].UniqueId;
                    if (data["selectedGridData"][index]["fileType"] == "file") {
                        this.spinnerService.show();
                        this.download$ = new Subject<any>();
                        // Get the SharePoint Sub Site if any

                        let subSite: string = this.utilityService.getSharePointSubSite(connector.absolute_url);
                        //get the cleared path
                        uploadPath = this.utilityService.clearSubSiteFromPath(subSite, relativePath);
                        //Replace file name from the upload path
                        uploadPath = uploadPath.replace(fileName, "");
                        //Shrujan download
                        //this.download(uploadPath, fileName, connector, entityConfiguration);
                        this.downloadSPAngular(uploadPath, fileName, connector, entityConfiguration, flag, fileUniqueId);
                        this.downloadSPSub = this.download$.subscribe(
                            (response) => {
                                if (response != "" && response != null && response["status"] == "true") {
                                    if (response["flag"] == true) {
                                        download(response["arrayBuffer"], response["fileName"]);
                                        self.modalService.fileDownloadingPercentage = "";
                                        self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['downloadingfile'] + " " + runningCount + "/" + count;
                                        if (count == runningCount) {
                                            self.spinnerService.hide();
                                            self.modalService.UploadStatusModalRef.hide();
                                            runningCount = 0;
                                        }
                                        runningCount++;
                                        //self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['downloadingfile'] + " " + runningCount + "/" + count;
                                    }
                                    else {
                                        fileExtension = response['fileName'].substring(response['fileName'].lastIndexOf(".") + 1, response['fileName'].length);
                                        blocked = this.utilityService.checkExtensionExistInCRM(fileExtension);
                                        if (!blocked) {
                                            //creating array for adding attachtments
                                            let fileColl: any = {};
                                            fileColl['base64'] = this.utilityService.arrayBufferToBase64(response['arrayBuffer']);
                                            fileColl['fileName'] = response['fileName'];
                                            selectedColl.push(fileColl);
                                            if ((count - countOfIgnore) == runningCount) {
                                                this.createEmailWithDocuments(selectedColl, entityConfiguration);
                                                self.spinnerService.hide();
                                            }
                                            runningCount++;
                                        }
                                        else {
                                            countOfIgnore++;
                                            if ((count == (data["selectedGridData"].length - countOfIgnore)) || ((data["selectedGridData"].length - countOfIgnore) == 0)) {
                                                this.spinnerService.hide();
                                            }
                                        }
                                    }
                                }
                                else {
                                    if (count == runningCount) {
                                        if (self.modalService.isOpen == false) {
                                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                        }
                                        self.modalService.UploadStatusModalRef.hide();
                                        self.spinnerService.hide();
                                    }
                                    runningCount++;
                                }
                            },
                            (error) => {
                                if (self.modalService.isOpen == false) {
                                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                }
                                self.utilityService.throwError(error, functionName);
                            }
                        );
                        // this.download(uploadPath, fileName, connector, entityConfiguration);
                    }
                    else {
                        runningCount++;
                        self.spinnerService.hide();
                    }
                }
            } else {
                this.spinnerService.show();
                fileName = path.split("/").pop();
                // Get the SharePoint Sub Site if any
                let subSite: string = this.utilityService.getSharePointSubSite(connector.absolute_url);
                //get the cleared path
                uploadPath = this.utilityService.clearSubSiteFromPath(subSite, path);
                //Replace file name from the upload path
                uploadPath = uploadPath.replace(fileName, "");
                this.download(uploadPath, fileName, connector, entityConfiguration);
                this.downloadSPSub = this.download$.subscribe(
                    (response) => {
                        if (response != "" && response != null) {
                            if (flag == true) {
                                download(response["base64"], response["fileName"]);
                                self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['downloadingfile'] + " " + runningCount + "/" + count;
                                if (count == runningCount) {
                                    self.spinnerService.hide();
                                    self.modalService.UploadStatusModalRef.hide();
                                    runningCount = 0;
                                }
                                runningCount++;
                                self.modalService.displayMessage = self.a2dAppService.labelsMultiLanguage['downloadingfile'] + " " + runningCount + "/" + count;
                            }
                            else {
                                fileExtension = response['fileName'].substring(response['fileName'].lastIndexOf(".") + 1, response['fileName'].length);
                                blocked = this.utilityService.checkExtensionExistInCRM(fileExtension);
                                if (!blocked) {
                                    //creating array for adding attachtments
                                    let fileColl: any = {};
                                    fileColl['base64'] = this.utilityService.arrayBufferToBase64(response['base64']);
                                    fileColl['fileName'] = response['fileName'];
                                    selectedColl.push(fileColl);
                                    if ((count - countOfIgnore) == runningCount) {
                                        this.createEmailWithDocuments(selectedColl, entityConfiguration);
                                        self.spinnerService.hide();
                                    }
                                    runningCount++;
                                }
                                else {
                                    countOfIgnore++;
                                }
                            }
                        }
                        else {
                            if (count == runningCount) {
                                self.spinnerService.hide();
                                if (self.modalService.isOpen == false) {
                                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                }
                            }
                            runningCount++;
                        }
                    },
                    (error) => {
                        if (self.modalService.isOpen == false) {
                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                        self.utilityService.throwError(error, functionName);
                    }
                );
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    downloadFile(url: string): any {
        const headers = new HttpHeaders({
            'Content-Type': 'application/octet-stream',
            'responseType': 'blob'
        });
        return this.http.get(url, { headers, responseType: 'arraybuffer' }).subscribe(
            (response: any) => {

            },
            (error) => {

            })
    }

    downloadSPAngular(relative_url: string, fileName: string, connector: Connector, entityConfiguration: EntityConfiguration, flag: boolean, fileUniqueId: string): any {
        let functionName: string = "downloadSPAngular";
        let self = this;
        let absoluteUrl: any = "";
        let relativePath: any = ""
        let newFileName: any = "";
        let odataQuery: any = "";
        let httpOptions = {};
        let decryptedToken: any = "";
        let downloadContent = {};
        let firstCharOfPath: any = "";
        let folderServerRelativeURL: any = "";
        let orignalRelativepath: string;
        try {
            orignalRelativepath = relative_url;
            absoluteUrl = connector.absolute_url;
            relativePath = relative_url.replace(/'/g, "''");
            newFileName = fileName.replace(/'/g, "''");
            //newFileName = this.a2dAppService.removeAllSpecialChar(newFileName,"-");
            firstCharOfPath = relative_url.charAt(0);
            if (firstCharOfPath == "/") {
                //path.substring(1, path.length)
                folderServerRelativeURL = encodeURIComponent(relativePath.substring(1, relativePath.length));
            }
            else {
                folderServerRelativeURL = encodeURIComponent(relativePath);
            }
            // odataQuery = `/_api/web/getFolderByServerRelativeUrl('${folderServerRelativeURL}')/files('${newFileName}')/$value?binaryStringResponseBody=true`;

            // When function hits for download file then use OpenBinaryStream() else use existing.
            if (flag == true) {
                odataQuery = `_api/web/getFileById('${fileUniqueId}')/OpenBinaryStream()`;
            }
            else {
                odataQuery = `_api/web/getFileById('${fileUniqueId}')/$value?binaryStringResponseBody=true`;
            }


            //@ts-ignore
            decryptedToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;
            let requestUrl = absoluteUrl + "/" + odataQuery;

            const headers = new HttpHeaders({
                'Content-Type': 'application/octet-stream',
                "Authorization": "Bearer " + decryptedToken
            });
            self.http.get(requestUrl, {
                reportProgress: true,
                observe: 'events',
                headers: headers,
                responseType: 'arraybuffer'
            }).subscribe
                (response => {
                    if (response.type === HttpEventType.DownloadProgress) {
                        // Calculate and display the percentage of upload completion
                        const percentDone = Math.round((100 * response.loaded) / response.total);
                        // console.log("Downloaded " + percentDone + "%");
                        self.modalService.fileDownloadingPercentage = `${percentDone}% Downloaded`;

                    }
                    else if (response instanceof HttpResponse) {
                        if (response.status == 200) {
                            let downloadedArrayBuffer: any = response.body;
                            // Use the base64 string as needed
                            if (this.a2dAppService.isValid(downloadedArrayBuffer)) {
                                self.modalService.fileDownloadingPercentage = "";
                                downloadContent["arrayBuffer"] = downloadedArrayBuffer;
                                downloadContent["fileName"] = fileName;
                                downloadContent["status"] = "true";
                                downloadContent["flag"] = flag;
                                self.download$.next(downloadContent);
                            }
                            else {
                                console.log("Not valid base64");
                                downloadContent["status"] = "false";
                                self.download$.next(downloadContent);
                            }
                        }
                        else {
                            self.a2dAppService.logError("", response.statusText, entityConfiguration, "", fileName, "");
                            downloadContent["status"] = "false";
                            self.download$.next(downloadContent);
                        }
                    }
                },
                    async (error) => {
                        debugger;
                        //this.a2dAppService.isValid(error.error.error_description) && (error.error.error_description = "Invalid JWT token. The token is expired.")
                        if (this.a2dAppService.isValid(error.status) && error.status == 401) {
                            const tokenResponse: any = await this.generateAccessTokensFromRefreshToken(connector, entityConfiguration);
                            this.acceessToken = tokenResponse.access_token;
                            if (tokenResponse != null && this.a2dAppService.isValid(tokenResponse.access_token)) {
                                //connector.access_token=tokenResponse.access_token;
                                this.downloadSPAngular(orignalRelativepath, fileName, connector, entityConfiguration, flag, fileUniqueId);

                            }
                            else {
                                downloadContent["status"] = "false";
                                self.download$.next(downloadContent);
                            }

                        }
                        else {
                            downloadContent["status"] = "false";
                            self.download$.next(downloadContent);
                        }
                    }
                );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }


    //Shrujan
    downloadSPAngularss(relative_url: string, fileName: string, connector: Connector, entityConfiguration: EntityConfiguration, flag: boolean): any {
        let functionName: string = "downloadSPAngular";
        let self = this;
        let absoluteUrl: any = "";
        let relativePath: any = ""
        let newFileName: any = "";
        let odataQuery: any = "";
        let httpOptions = {};
        let decryptedToken: any = "";
        let downloadContent = {};
        let firstCharOfPath: any = "";
        let folderServerRelativeURL: any = "";
        try {
            absoluteUrl = connector.absolute_url;
            relativePath = relative_url.replace(/'/g, "''");
            newFileName = fileName.replace(/'/g, "''");
            //newFileName = this.a2dAppService.removeAllSpecialChar(newFileName,"-");
            firstCharOfPath = relative_url.charAt(0);
            if (firstCharOfPath == "/") {
                //path.substring(1, path.length)
                folderServerRelativeURL = encodeURIComponent(relativePath.substring(1, relativePath.length));
            }
            else {
                folderServerRelativeURL = encodeURIComponent(relativePath);
            }

            odataQuery = `/_api/web/getFolderByServerRelativeUrl('${folderServerRelativeURL}')/files('${newFileName}')/$value?binaryStringResponseBody=true`;
            //@ts-ignore
            decryptedToken = InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;
            let requestUrl = absoluteUrl + "/" + odataQuery;
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + decryptedToken
                }),
                responseType: 'blob'
            };
            self.http.get(requestUrl, httpOptions).subscribe(
                (response: any) => {
                    if (response.size > 0) {
                        let reader: any = new FileReader();
                        reader.onloadend = () => {
                            let downloadedbase64: any = btoa(reader.result as string);
                            // Use the base64 string as needed
                            if (this.a2dAppService.isValid(downloadedbase64)) {
                                //delete response["status"]
                                downloadContent["base64"] = downloadedbase64;
                                downloadContent["fileName"] = fileName;
                                downloadContent["status"] = "true";
                                downloadContent["flag"] = flag;
                                // this.downloadSPFileCount = this.downloadSPFileCount + 1;
                                self.download$.next(downloadContent);
                            }
                            else {
                                console.log("Not valid base64");
                                downloadContent["status"] = "false";
                                self.download$.next(downloadContent);
                            }
                        };
                        reader.readAsBinaryString(response);
                    }
                    else {
                        self.spinnerService.hide();
                        if (self.modalService.isOpen == false) {
                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                    }
                },
                (error) => {
                    self.spinnerService.hide();
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * create email and attach files in emails
     * @param files
     */
    createEmailWithDocuments(files: any, entityConfiguration: any) {

        let functionName: string = "createEmailWithDocuments: ";
        let self: any = this;
        let email: any = {};
        let parties = [];
        var fromSendParty = {};
        var toSendParty = {};
        var ccSendParty = {};
        var bccSendParty = {};

        this._Xrm = new IKL.Attach2Dynamics.CrmJs();

        //parameters = this.a2dAppService.getUCIQueryStringParameters();
        let emaiId: string;
        let data: any;
        let attachmentAddeed = this.a2dAppService.labelsMultiLanguage['success'];
        let okMessage: string = this.a2dAppService.labelsMultiLanguage['tooltipok'];
        let alertStrings = { confirmButtonLabel: okMessage, text: attachmentAddeed, title: "" };
        let alertOptions = { height: 130, width: 140 };
        try {

            // Get logged in user id.
            let userId = this.a2dAppService._Xrm.getUserId().substring(1, this.a2dAppService._Xrm.getUserId().length - 1);
            let queryOptions = {
                includeFormattedValues: true,
                includeLookupLogicalNames: true,
                includeAssociatedNavigationProperties: true
            }

            self.retrieveRecord$ = new Subject<any>();
            // Retrieve record to set dynamic values in email recipients.
            self.a2dAppService.webApi.retrieve("" + this.a2dAppService.currentEntityPluralName + "", this.a2dAppService.currentEntityId, "", queryOptions).then(
                (response) => {
                    response = this.a2dAppService.extractResponse(response);

                    var entityMetaData = "EntityDefinitions?$select=EntitySetName,LogicalName";
                    self.a2dAppService.webApi.retrieveMultiple(entityMetaData, null, null).then(function (metadataResponse) {
                        metadataResponse = self.a2dAppService.extractResponse(metadataResponse);
                        //store entityset name and primary attribute 
                        // _this.currentEntityPluralName = _this.isValid(response[0]) ? response[0].EntitySetName : "";
                        // #Added 16/12/19
                        //this._queryStringParamObjData["EntitySetName"] = this.currentEntityPluralName;

                        //ActivityParty (From)
                        if (self.a2dAppService.isValid(entityConfiguration) && self.a2dAppService.isValid(entityConfiguration["fromSender"]) && self.a2dAppService.isValid(JSON.parse(entityConfiguration["fromSender"])[0])) {
                            // Get value of from stored in email configuration.
                            fromSendParty = self.a2dAppService.getEmailSender(response, entityConfiguration, userId);
                        } else {
                            fromSendParty["partyid_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
                        }
                        fromSendParty["participationtypemask"] = 1; //From
                        parties.push(fromSendParty);

                        //ActivityParty (To)
                        if (self.a2dAppService.isValid(entityConfiguration) && self.a2dAppService.isValid(entityConfiguration["toRecipient"])) {
                            let toRecipients = JSON.parse(entityConfiguration["toRecipient"]);
                            for (let index = 0; index < toRecipients.length; index++) {
                                let toRecipientData = toRecipients[index];
                                toSendParty = self.a2dAppService.getEmailRecipients(metadataResponse, response, entityConfiguration, userId, toRecipientData, 2, parties)
                                if (self.a2dAppService.isValid(toSendParty) && toSendParty.hasOwnProperty("participationtypemask") && (toRecipientData["entitytype"] != "team" && toRecipientData["entitytype"] != "manager")) {
                                    parties.push(toSendParty);
                                }
                            }
                        }
                        else if (self.a2dAppService.isActitvityParty == true) {
                            toSendParty["partyid_" + self.a2dAppService.currentEntityName + "@odata.bind"] = "/" + self.a2dAppService.currentEntityPluralName + "(" + self.a2dAppService.currentEntityId + ")";
                            toSendParty["participationtypemask"] = 2; //To
                            parties.push(toSendParty);
                        }

                        //ActivityParty (Cc)
                        if (self.a2dAppService.isValid(entityConfiguration) && self.a2dAppService.isValid(entityConfiguration["ccRecipient"])) {
                            let ccRecipient = JSON.parse(entityConfiguration["ccRecipient"]);
                            for (let index = 0; index < ccRecipient.length; index++) {
                                let ccRecipientData = ccRecipient[index];
                                ccSendParty = self.a2dAppService.getEmailRecipients(metadataResponse, response, entityConfiguration, userId, ccRecipientData, 3, parties)
                                if (self.a2dAppService.isValid(ccSendParty) && ccSendParty.hasOwnProperty("participationtypemask") && (ccRecipientData["entitytype"] != "team" && ccRecipientData["entitytype"] != "manager")) {
                                    parties.push(ccSendParty);
                                }
                            }
                        }

                        //ActivityParty (Bcc)
                        if (self.a2dAppService.isValid(entityConfiguration) && self.a2dAppService.isValid(entityConfiguration["bccRecipient"])) {
                            let bccRecipient = JSON.parse(entityConfiguration["bccRecipient"]);
                            for (let index = 0; index < bccRecipient.length; index++) {
                                let bccRecipientData = bccRecipient[index];
                                bccSendParty = self.a2dAppService.getEmailRecipients(metadataResponse, response, entityConfiguration, userId, bccRecipientData, 4, parties)
                                if (self.a2dAppService.isValid(bccSendParty) && bccSendParty.hasOwnProperty("participationtypemask") && (bccRecipientData["entitytype"] != "team" && bccRecipientData["entitytype"] != "manager")) {
                                    parties.push(bccSendParty);
                                }
                            }
                        }

                        email["email_activity_parties"] = parties;
                        //Regarding
                        if (self.a2dAppService.isActitvityParty == true) {
                            let regarding = "/" + self.a2dAppService.currentEntityPluralName + "(" + self.a2dAppService.currentEntityId + ")";
                            let dynamicsLookup = "regardingobjectid_" + self.a2dAppService.currentEntityName + "@odata.bind";
                            email[dynamicsLookup] = regarding;
                        }

                        //Condition to check if A2D UI is open from Email form or not.
                        if ((self.a2dAppService['actualCurrentEntity'] && self.a2dAppService['actualCurrentEntityId'])) {
                            //if opened from email form follow the below actions.

                            self.a2dAppService.webApi.retrieve("" + self.a2dAppService.actualCurrentEntity + "", self.a2dAppService.actualCurrentEntityId, "", queryOptions).then(
                                (response) => {
                                    response = self.a2dAppService.extractResponse(response);

                                    self.a2dAppService.bindDocuments(response.activityid, files);
                                    /*
                                        self.a2dAppService.message_success = 
                                        self.modalService.isOpen = true;
                                        self.modalService.openAlertDialog(self.a2dAppService.message_success, (onOKClick) => { })
                                        */
                                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                                        function (success) {

                                        },
                                        function (error) {
                                            console.log(error.message);
                                        }
                                    );
                                },
                                (error: any) => {
                                    console.log(error.message);
                                });
                            // self.modalService.isOpen = true;
                            // self.modalService.openAlertDialog(this.a2dAppService.message_commonError, (onOKClick) => { })
                        }
                        else {
                            self.a2dAppService.webApi.create("emails", email).then((response: any) => {
                                response = self.a2dAppService.extractResponse(response);

                                self.a2dAppService.bindDocuments(response.id.value, files);
                            }, (error: any) => {
                                let isAppendPrivError = JSON.parse(error.response.data).error.message.includes("AppendToAccess");
                                if (isAppendPrivError) {
                                    self.createErrorLog(JSON.parse(error.response.data).error.message, entityConfiguration, "");
                                    self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                                }
                                self.utilityService.throwError(error, functionName);
                            });
                        }

                    }, function (error) {
                        console.log(functionName + "> retrieveMultiple: " + error.message || error.description);
                    });

                },
                (error) => {
                    self.utilityService.throwError(error, functionName);
                }
            );

        } catch (error) {
            this.a2dAppService.logError('', error.description || error.message, entityConfiguration, '', null, null);
            this.modalService.isOpen = true;
            this.modalService.openErrorDialog(this.a2dAppService.message_commonError, (onOKClick) => { });

            this.utilityService.throwError(error, functionName);

        }
    }
    /**
     * Download Files from SP
     * @param relative_url
     * @param fileName
     * @param connector
     * @param entityConfiguration
     */
    download(relative_url: string, fileName: string, connector: Connector, entityConfiguration: EntityConfiguration): void {
        let functionName: string = "download";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        try {
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            let additionalDetail: {} = {};
            additionalDetail["relative_path"] = relative_url;
            additionalDetail["file_name"] = fileName;
            fileDetail.file_name = fileName;
            fileDetail.path = relative_url;
            let object = {
                "MethodName": "downloadfile",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "AdditionalDetailsJSON": JSON.stringify(additionalDetail),
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration),

            }
            this.download$ = new Subject<any>();
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "Download");
            this.downloadSub = this.a2dAppService.download$.subscribe(
                (response) => {
                    if (response["status"]) {
                        delete response["status"]
                        self.download$.next(response);
                    }
                    else {
                        self.download$.next("");
                    }
                },
                (error) => {

                    self.download$.next("");
                    self.utilityService.throwError(error, functionName);
                }
            );
            // this.a2dAppService.callSharePointCoreAction(object, this.actionName, "Download");
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * View file in SharePoint
     * @param uploadPath
     * @param newName
     * @param oldName
     * @param connector
     * @param entityConfiguration
     */
    viewFile(uploadPath: string, newName: string, oldName: string, connector: Connector, entityConfiguration: EntityConfiguration): void {
        let functionName: string = "viewFile ";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        try {
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            let additionalDetail: {} = {};
            additionalDetail["relative_path"] = uploadPath;
            additionalDetail["new_name"] = newName;
            additionalDetail["old_name"] = oldName;
            fileDetail.file_name = newName;
            fileDetail.path = uploadPath;
            let object = {
                "MethodName": "viewfile",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "AdditionalDetailsJSON": JSON.stringify(additionalDetail),
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration)
            }
            this.a2dAppService.viewFileSP$ = new Subject<any>();
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "ViewFile");
            this.viewFileSub = this.a2dAppService.viewFileSP$.subscribe(
                (response) => {
                    this.viewFile$.next(response["link"]);
                },
                (error) => {

                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Rename file in SharePoint
     * @param uploadPath
     * @param newName
     * @param oldName
     * @param connector
     * @param entityConfiguration
     */
    renameFile(uploadPath: string, newName: string, oldName: string, connector: Connector, entityConfiguration: EntityConfiguration, UniqueId: string): void {
        let functionName: string = "renameFile ";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        try {
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            let additionalDetail: {} = {};
            additionalDetail["relative_path"] = uploadPath;
            additionalDetail["new_name"] = newName;
            additionalDetail["old_name"] = oldName;
            additionalDetail["uniqueId"] = UniqueId;
            additionalDetail["rootLibraryDisplayName"] = entityConfiguration.rootLibraryDisplayName;
            fileDetail.file_name = newName;
            fileDetail.path = uploadPath;
            // Rename the file in CRM
            let object = {
                "MethodName": "renamefile",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "AdditionalDetailsJSON": JSON.stringify(additionalDetail),
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration)
            }
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "RenameFile");
            this.renameFileSub = this.a2dAppService.renameFileSP$.subscribe(
                (response) => {
                    if (response["status"] == "true" || response["status"] == true) { }
                    else if (response["status"] == "false" || response["status"] == false) {
                        if (self.modalService.isOpen == false) {
                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                    }
                    self.getSharePointData(connector, entityConfiguration, uploadPath, self.gridService.selectedView);
                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
       * Delete file in SharePoint
       * @param uploadPath
       * @param newName
       * @param oldName
       * @param connector
       * @param entityConfiguration
       */
    deleteFile(uploadPath: string, newName: string, oldName: string, connector: Connector, entityConfiguration: EntityConfiguration, UniqueId: string, forMovedFile: boolean): void {
        let functionName: string = "deleteFile ";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        try {
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            let additionalDetail: {} = {};
            additionalDetail["relative_path"] = uploadPath;
            additionalDetail["uniqueId"] = UniqueId;
            additionalDetail["new_name"] = newName;
            additionalDetail["old_name"] = oldName;
            fileDetail.file_name = newName;
            fileDetail.path = uploadPath;
            // Delete the file in CRM
            let object = {
                "MethodName": "deletefile",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "AdditionalDetailsJSON": JSON.stringify(additionalDetail),
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration)
            }
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "DeleteFile");
            this.deleteFileSub = this.a2dAppService.deleteFileSP$.subscribe(
                (response) => {
                    if (response["status"] == "true" || response["status"] == true) {
                        if (!forMovedFile) {
                            this.a2dAppService.logError('', "Deleted Permanently.", entityConfiguration, '', fileDetail.file_name, uploadPath, "delete");
                        }
                    }
                    else if (response["status"] == "false" || response["status"] == false) {
                        if (self.modalService.isOpen == false) {
                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                    }
                    if (forMovedFile == false) {
                        self.getSharePointData(connector, entityConfiguration[0], uploadPath, self.gridService.selectedView);
                    }


                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }


    /**
    * Rename folder in SharePoint
    * @param uploadPath
    * @param newName
    * @param oldName
    * @param connector
    * @param entityConfiguration
    */
    renameFolder(uploadPath: string, newName: string, oldName: string, connector: Connector, entityConfiguration: EntityConfiguration, UniqueId: string): void {
        let functionName: string = "renameFolder ";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        try {
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            let additionalDetail: {} = {};
            additionalDetail["relative_path"] = uploadPath;
            additionalDetail["new_name"] = newName;
            additionalDetail["old_name"] = oldName;
            additionalDetail["uniqueId"] = UniqueId;
            fileDetail.file_name = newName;
            fileDetail.path = uploadPath;
            // Rename the file in CRM
            let object = {
                "MethodName": "renamefolder",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "AdditionalDetailsJSON": JSON.stringify(additionalDetail),
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration)
            }
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "RenameFolder");
            this.renameFolderSub = this.a2dAppService.renameFolderSP$.subscribe(
                (response) => {
                    if (response["status"] == "true" || response["status"] == true) { }
                    else if (response["status"] == "false" || response["status"] == false) {
                        if (self.modalService.isOpen == false) {
                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                    }
                    self.getSharePointData(connector, entityConfiguration, uploadPath, self.gridService.selectedView);
                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
    * check whether it is other than web.
    */
    getDevice(): boolean {
        let functionName: string = "getDevice";
        let isItWeb: boolean = true;
        try {
            if (/Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent)) {
                isItWeb = false;
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return isItWeb;
    }

    /**
     * Create a Share Link
     * @param path
     * @param isEditLink
     * @param connector
     * @param entityConfiguration
     * @param fileName
     */
    shareALink(path: string, isEditLink: boolean, connector: Connector, entityConfiguration: EntityConfiguration, fileName: string, uniqueId: string, fileType: string, flag: boolean, acesslinkJson?: string): void {
        let functionName: string = "shareALink";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        let fileDetail: FileDetail = {};
        try {
            // Get the SharePoint Sub Site if any
            let subSite: string = this.utilityService.getSharePointSubSite(connector.absolute_url);
            //get the cleared path
            path = this.utilityService.clearSubSiteFromPath(subSite, path);
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            this.shareLink$ = new Subject<any>();
            let additionalDetail: {} = {};
            additionalDetail["isEditLink"] = isEditLink;
            additionalDetail["url"] = path;
            additionalDetail["file_name"] = fileName;
            additionalDetail["uniqueId"] = uniqueId;
            additionalDetail["fileType"] = fileType;
            additionalDetail["acesslinkJson"] = acesslinkJson;
            fileDetail.file_name = fileName;
            fileDetail.path = path;
            // Share File Link in CRM
            let object = {
                "MethodName": "sharelink",
                "ConnectorJSON": JSON.stringify(connector),
                "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
                "EntityName": entityName,
                "RecordId": recordId,
                "AdditionalDetailsJSON": JSON.stringify(additionalDetail),
                "FileDetailsJSON": JSON.stringify(fileDetail),
                "EntityConfigurationJSON": JSON.stringify(entityConfiguration)
            }
            this.a2dAppService.callSharePointCoreAction(object, this.actionName, "ShareLink");
            this.shareLinkSub = this.a2dAppService.shareLinkSP$.subscribe(
                (response) => {
                    this.spinnerService.hide();
                    if (response && response["status"] && response["status"] == "true" || response["status"] == true) {
                        self.modalService.inputValue = response["link"];
                        if (flag == true) {
                            self.modalService.openDialogWithInputShareLink(this.a2dAppService.labelsMultiLanguage['sharedlink'] + " " + fileName,
                                (onOKClick) => {
                                    this.utilityService.copyText(response["link"]);
                                }
                            );
                        }
                        else {
                            let fileUrlColl: {} = {};
                            fileUrlColl["name"] = response['fileName'];
                            fileUrlColl["url"] = response['link'];
                            this.shareLink$.next(fileUrlColl);
                        }
                    }
                    else {
                        if (self.modalService.isOpen == false) {
                            self.modalService.isOpen = true;
                            self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                        }
                    }
                },
                (error) => {
                    if (self.modalService.isOpen == false) {
                        self.modalService.isOpen = true;
                        self.modalService.openErrorDialog(self.a2dAppService.message_commonError, (onOKClick) => { });
                    }
                    // self.utilityService.throwError(error, functionName);
                }
            );
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Search files/folders based on the text
     * @param relative_path
     * @param searchWords
     * @param list
     * @param connector
     * @param entityConfiguration
     */
    searchFiles(relative_path: string, searchWords: string, list: string, connector: Connector, entityConfiguration: EntityConfiguration) {
        let functionName: string = "searchFiles";
        let parameters: any = null;
        let entityName: string = "";
        let recordId: string = "";
        let self = this;
        try {
            entityName = this.a2dAppService.currentEntityName;
            recordId = this.a2dAppService.currentEntityId;
            this.shareLink$ = new Subject<any>();
            this.getFilesAngular$ = new Subject<any>();
            // let sharePointSubSite: string = this.utilityService.getSharePointSubSite(connector.absolute_url);
            // sharePointSubSite = sharePointSubSite && sharePointSubSite.endsWith("/") ?
            //     sharePointSubSite.substr(0, sharePointSubSite.length - 1) : sharePointSubSite;
            // list = sharePointSubSite ? `${sharePointSubSite}/${list}` : list;
            // if (this.utilityService.isValid(sharePointSubSite)) {
            //     // relative_path = relative_path.startsWith(sharePointSubSite) ? relative_path.replace(`${sharePointSubSite}/`, '') : relative_path;
            //     // relative_path = relative_path.startsWith(`/${sharePointSubSite}`) ? relative_path.replace(`/${sharePointSubSite}/`, '') : relative_path;
            //     // relative_path = sharePointSubSite ? `/${sharePointSubSite}/${relative_path}` : relative_path;
            //     //get the cleared path
            //     relative_path = this.utilityService.clearSubSiteFromPath(sharePointSubSite, relative_path);
            //     relative_path = relative_path && relative_path.startsWith("/") ?
            //         relative_path.substr(1, relative_path.length) : relative_path;
            //     relative_path = sharePointSubSite ? `/${sharePointSubSite}/${relative_path}` : relative_path;
            // }
            // let additionalDetail: {} = {};
            // additionalDetail["relative_path"] = relative_path;
            // additionalDetail["list"] = list;
            // additionalDetail["search_word"] = searchWords;
            // // Search in SharePoint
            // let object = {
            //     "MethodName": "search",
            //     "ConnectorJSON": JSON.stringify(connector),
            //     "CRMURL": this.a2dAppService._Xrm.getClientUrl(),
            //     "EntityName": entityName,
            //     "RecordId": recordId,
            //     "AdditionalDetailsJSON": JSON.stringify(additionalDetail),
            //     "EntityConfigurationJSON": JSON.stringify(entityConfiguration)
            // }
            // this.a2dAppService.callSharePointCoreAction(object, this.actionName, "Search");
            // this.searchSub = this.a2dAppService.searchContentSP$.subscribe(
            //     (response) => {
            //         if (response.status == "true" || response.status == true) {
            //             let files: GridData[] = self.createCollectionSharePointData(response["Files"], null, connector, entityConfiguration);
            //             self.a2dAppService.showEmptyDataMessage = files.length > 0 ? false : true;
            //             self.search$.next(files);
            //         }
            //         else {
            //             self.a2dAppService.showEmptyDataMessage = true;
            //             self.search$.next([{}]);
            //         }
            //     },
            //     (error) => {
            //         self.a2dAppService.showEmptyDataMessage = true;
            //         self.search$.next([{}]);
            //         self.utilityService.throwError(error, functionName);
            //     }
            // );
            let selectedView: any;
            if (this.a2dAppService.isValid(this.gridService.selectedView) && this.gridService.selectedView != " Thumbnail View ") {
                selectedView = self.a2dAppService.Views.find(view => view.Id == this.gridService.selectedView)
            }
            else if (this.a2dAppService.isValid(this.gridService.selectedViewThumbnail)) {
                selectedView = self.a2dAppService.Views.find(view => view.Id == this.gridService.selectedViewThumbnail)
            }
            else {
                selectedView = self.a2dAppService.Views.find(view => view.DefaultView == true);
            }
            let viewXml = this.updateView(selectedView.ListViewXml, this.colFields, false)
            this.getFilesAngular(connector, entityConfiguration[0], relative_path, "", viewXml, "SearchFiles", "", searchWords);
            this.getFilesAngularSub = this.getFilesAngular$.subscribe((getFilesResponse: any) => {
                let files: any = self.createCollectionSharePointData(getFilesResponse.ListData.Row, self.a2dAppService.columnsArray, connector, entityConfiguration);
                //self.spinnerService.hide();
                self.search$.next(files);
            });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Get the base64 from the reader.result value
     * @param base64
     */
    getBase64(base64: any): string {
        let functionName: string = "getBase64";
        let formattedBase64: string = "";
        try {
            formattedBase64 = base64.substring(base64.indexOf("base64,") + 7, base64.length);
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return formattedBase64;
    }

    /**
     * Create uploaded file details object, this gives the object to show the details in the final status UI
     * @param response
     */
    createUploadedFileDetailsObject(response): string[] {
        let functionName: string = "createUploadedFileDetailsObject";
        let uploadedFileDetails: any = [];
        try {
            uploadedFileDetails["FileName"] = response["FileName"];
            uploadedFileDetails["FilePath"] = response["FilePath"];
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return uploadedFileDetails;
    }

    // Shrujan -  Thumnails view
    getSPThumbnails(connector: Connector, selectedEntityConfiguration: EntityConfiguration, fileData: any) {
        let functionName: string = "getThumbnails";
        let sharePointSite: any;
        let thumbnailURL: any;
        let pathUrl: any;
        let previewPath: any;
        try {
            if (this.a2dAppService.isValid(connector)) {
                sharePointSite = this.utilityService.getSharePointSite(connector.absolute_url);
                // When UI is opning from homegrid with multiple records. 
                if (this.a2dAppService.isValid(selectedEntityConfiguration) && !this.a2dAppService.isValid(fileData.path_display)) {
                    pathUrl = connector.absolute_url + "/" + selectedEntityConfiguration[0].folder_path + "/" + fileData.fileName;
                    previewPath = pathUrl.replace(/([^:]\/)\/+/g, "$1");// Using url regex for accurate URL
                    thumbnailURL = sharePointSite + "/_layouts/15/getpreview.ashx?path=" + previewPath + "&resolution=3&force=1";
                }
                else { //When UI is opning from single records.
                    previewPath = sharePointSite + fileData.path_display;
                    thumbnailURL = sharePointSite + "/_layouts/15/getpreview.ashx?path=" + previewPath + "&resolution=3&force=1";
                }
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return thumbnailURL;
    }


    /**
    * This function performs destroy of subscriptions
    */
    ngOnDestroy() {
        this.destroySubscriptions();
    }
    //#region Added on 06/09/2024 By Lakshman for Lookup and person on sharepoint
    App_AddLookUpListItems(connector: Connector, entityConfiguration: EntityConfiguration, recordId: string, displayName: string, lookUpEntityName: string, listId: string, sharePointColumn: any, isColumnPresent: boolean) {
        let functionName: string = "App_AddLookUpListItems: ";
        // let displayName: string = "";
        // let recordId: string = "";
        let self = this;
        let requestUrl: string;
        let httpOptions: any;
        let lookUpId: any;
        let listName: any;
        // const fieldValueMap: { [key: string]: { type: string, value: any } } = {};
        const fieldValueMap: { [key: string]: string } = {};
        try {
            self = this;

            if (this.a2dAppService.isValid(isColumnPresent) && isColumnPresent == true) {
                requestUrl = connector.absolute_url + "/_api/web/Lists(guid'" + listId + "')/items?$filter=RecordId eq '" + recordId + "' or Title eq '" + encodeURIComponent(displayName) + "'";
            }
            else {
                requestUrl = connector.absolute_url + "/_api/web/Lists(guid'" + listId + "')/items?$filter=Title eq '" + encodeURIComponent(displayName) + "'";
            }
            //@ts-ignore
            this.acceessToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + this.acceessToken,
                    "Content-Type": "application/json;odata=verbose"
                }),
            };
            this.http.get(requestUrl, httpOptions,).subscribe((resp: any) => {
                //Validate the result and get the id of lookup or create new item in list
                if (self.a2dAppService.isValid(resp) && self.a2dAppService.isValid(resp.d) && self.a2dAppService.isValid(resp.d.results) && self.a2dAppService.isValid(resp.d.results[0])) {
                    fieldValueMap['key'] = sharePointColumn;
                    fieldValueMap['value'] = resp.d.results[0].ID.toString();
                    //fieldValueMap[sharePointColumn] = { type: sharePointColumn, value: resp.d.results[0].ID };
                    // lookUpId=resp.d.results[0].ID;
                    self.App_AddLookUpListItems$.next(fieldValueMap);
                }
                else {
                    self.App_GetListName$ = new Subject<any>()
                    self.App_GetListName(connector, entityConfiguration, listId, recordId, displayName, sharePointColumn, lookUpEntityName, isColumnPresent);
                    self.App_GetListNameSub = self.App_GetListName$.subscribe((response: any) => {
                        self.App_AddLookUpListItems$.next(response);
                    });
                }

            }, (err) => {
                if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Invalid JWT token. The token is expired.")) {
                    self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then((tokenResponse: any) => {
                        self.acceessToken = tokenResponse ? tokenResponse.access_token : null;
                        if (tokenResponse && self.a2dAppService.isValid(tokenResponse.access_token)) {
                            self.App_AddLookUpListItems(connector, entityConfiguration, recordId, displayName, lookUpEntityName, listId, sharePointColumn, true);
                        }
                    })
                }
                else if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Column 'RecordId' does not exist")) {
                    self.App_AddLookUpListItems(connector, entityConfiguration, recordId, displayName, lookUpEntityName, listId, sharePointColumn, false);
                }
            });

        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    App_GetListName(connector: Connector, entityConfiguration: EntityConfiguration, listId: string, recordId: string, displayName: string, sharePointColumn: any, entityName: string, isColumnPresent: boolean) {
        let functionName: string = "App_GetListName";
        let requestUrl: any;
        let httpOptions: any;
        let self: any;
        let listName: any;
        let data: any
        const fieldValueMap: { [key: string]: string } = {};
        try {
            self = this;
            requestUrl = connector.absolute_url + "/_api/web/Lists(guid'" + listId + "')?$select=ListItemEntityTypeFullName'"
            //@ts-ignore
            this.acceessToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + this.acceessToken,
                    "Content-Type": "application/json;odata=verbose"
                }),
            };
            this.http.get(requestUrl, httpOptions,).subscribe((resp: any) => {
                //Validate the result and get the id of lookup or create new item in list
                if (self.a2dAppService.isValid(resp) && self.a2dAppService.isValid(resp.d) && self.a2dAppService.isValid(resp.d.ListItemEntityTypeFullName)) {
                    listName = resp.d.ListItemEntityTypeFullName;
                    //#region 
                    requestUrl = connector.absolute_url + "/_api/web/Lists(guid'" + listId + "')/items";
                    // httpOptions = {
                    //     headers: new HttpHeaders({
                    //         "accept": "application/json;odata=verbose",
                    //         "Authorization": "Bearer " + this.acceessToken,
                    //         "Content-Type": "application/json;odata=verbose"
                    //     }),
                    // };
                    if (isColumnPresent) {
                        //data = JSON.stringify({ "parameters": { "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": 7837447, "AllowMultipleValueFilterForTaxonomyFields": true, "AddRequiredFields": true, "RequireFolderColoringFields": true } }); //"ViewXml":"<RowLimit Paged=\"TRUE\">1000</RowLimit>", //,"ViewXml":"<View><RowLimit Paged=\"TRUE\">30</RowLimit><QueryOptions><Paging ListItemCollectionPositionNext=\"\"/></QueryOptions></View>"
                        data = JSON.stringify({ __metadata: { type: listName }, Title: displayName, RecordId: recordId });
                    }
                    else {
                        data = JSON.stringify({ __metadata: { type: listName }, Title: displayName });
                    }
                    self.executePost(requestUrl, data, httpOptions, connector, entityConfiguration).then(response => {
                        if (response.hasError == true) {

                        }
                        else {
                            if (self.a2dAppService.isValid(response) && self.a2dAppService.isValid(response.ID)) {
                                fieldValueMap['key'] = sharePointColumn;
                                fieldValueMap['value'] = response.ID.toString();
                                self.App_GetListName$.next(fieldValueMap);
                            }
                            else {
                                // lookUpId="";
                            }
                        }
                    }, (err) => {
                        if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Invalid JWT token. The token is expired.")) {
                            self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then((tokenResponse: any) => {
                                self.acceessToken = tokenResponse ? tokenResponse.access_token : null;
                                if (tokenResponse && self.a2dAppService.isValid(tokenResponse.access_token)) {
                                    self.App_GetListName(connector, entityConfiguration, listId, recordId, displayName, sharePointColumn, entityName, true);
                                }
                            })
                        }
                        else if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Column 'RecordId' does not exist")) {
                            self.App_GetListName(connector, entityConfiguration, listId, recordId, displayName, sharePointColumn, entityName, false);
                        }
                    });

                    //#endregion

                }
            }, (err) => {
                if (err && err.error && err.error.error_description && err.error.error_description.includes("Invalid JWT token. The token is expired.")) {
                    self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then((tokenResponse: any) => {
                        self.acceessToken = tokenResponse ? tokenResponse.access_token : null;
                        if (tokenResponse && self.a2dAppService.isValid(tokenResponse.access_token)) {
                            self.App_GetListName(connector, entityConfiguration, listId, recordId, displayName, sharePointColumn, entityName, isColumnPresent);
                        }
                    })
                }
                else if (err && err.error && err.error.error_description && err.error.error_description.includes("Column 'RecordId' does not exist")) {
                    self.App_GetListName(connector, entityConfiguration, listId, recordId, displayName, sharePointColumn, entityName, isColumnPresent);
                }
            });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    App_AddUserListItems(connector: Connector, entityConfiguration: EntityConfiguration, recordId: string, attributeName: string, sharePointColumn: string) {
        let functionName: string = "App_AddUserListItems:";
        // let fetchXml:string="";
        let self: any;
        // let recordId: any;
        let entityName: string = "systemuser";
        const fieldValueMap: { [key: string]: string } = {};
        let columnSet: any = ["domainname"];;
        try {
            self = this;
            //let matchingRecord:any=this.a2dAppService._recordValuesArray.find(record => record.entityName === parentEntityName.toLowerCase());
            // recordId = this.a2dAppService._recordValuesArray[0][0][`_${attributeName}_value`];
            Xrm.WebApi.retrieveRecord(entityName, recordId, `?$select=${columnSet.join(",")}`).then(userRecord => {
                if (userRecord != null && userRecord.domainname != null) {
                    fieldValueMap['key'] = sharePointColumn;
                    fieldValueMap['value'] = userRecord.domainname;
                    self.App_AddUserListItems$.next(fieldValueMap);
                }
            }).catch(error => {
                self.utilityService.throwError(error, functionName);
            });
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    //#endregion
    //#region 

    getSharePointLookUpValues(connector: Connector, entityConfiguration: EntityConfiguration, col: any, type: any) {
        let functionName: string = "getSharePointLookUpValues:";
        let requestUrl: any;
        let httpOptions: any;
        let dropdownItems: { label: string, value: string }[] = [];
        let self: any;
        try {
            self = this;
            //@ts-ignore
            this.acceessToken = this.a2dAppService.isValid(this.acceessToken) ? this.acceessToken : InoEncryptDecrypt.EncryptDecrypt.decryptKey(connector.access_token).DecryptedValue;
            httpOptions = {
                headers: new HttpHeaders({
                    "accept": "application/json;odata=verbose",
                    "Authorization": "Bearer " + this.acceessToken,
                    "Content-Type": "application/json;odata=verbose"
                }),
            };
            switch (type) {
                case "lookup":
                    requestUrl = connector.absolute_url + "/_api/web/Lists(guid'" + col.listId + "')/items?$select=Id,Title";
                    this.http.get(requestUrl, httpOptions,).subscribe((resp: any) => {
                        dropdownItems = resp.d.results.map((item: any) => ({
                            label: item.Title,  // The 'title' from your API response
                            value: item.Id      // The 'id' from your API response
                        }));
                        self.getSharePointLookUpValues$.next(dropdownItems);
                    }, (err) => {
                        if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Invalid JWT token. The token is expired.")) {
                            self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then((tokenResponse: any) => {
                                self.acceessToken = tokenResponse ? tokenResponse.access_token : null;
                                if (tokenResponse && self.a2dAppService.isValid(tokenResponse.access_token)) {
                                    self.getSharePointLookUpValues(connector, entityConfiguration, col, type);
                                }
                            })
                        }
                        else if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Column 'RecordId' does not exist")) {
                            self.getSharePointLookUpValues(connector, entityConfiguration, col, type);
                        }
                    });
                    break;
                case "user":
                    requestUrl = connector.absolute_url + "/_api/web/siteusers?$select=Email,Title&$filter=substringof('.com', Email)";
                    this.http.get(requestUrl, httpOptions,).subscribe((resp: any) => {
                        dropdownItems = resp.d.results.map((item: any) => ({
                            label: item.Title,  // The 'title' from your API response
                            value: "[{\"key\":\"i:0#.f|membership|" + item.Email + "\"}]"
                            // value: item.Email      // The 'id' from your API response
                        }));
                        self.getSharePointLookUpValues$.next(dropdownItems);
                    }, (err) => {
                        if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Invalid JWT token. The token is expired.")) {
                            self.generateAccessTokensFromRefreshToken(connector, entityConfiguration).then((tokenResponse: any) => {
                                self.acceessToken = tokenResponse ? tokenResponse.access_token : null;
                                if (tokenResponse && self.a2dAppService.isValid(tokenResponse.access_token)) {
                                    self.getSharePointLookUpValues(connector, entityConfiguration, col, type);
                                }
                            })
                        }
                        else if (err && err.error && err.error.error && err.error.error.message && err.error.error.message.value && err.error.error.message.value.includes("Column 'RecordId' does not exist")) {
                            self.getSharePointLookUpValues(connector, entityConfiguration, col, type);
                        }
                    });
                    break;
            }
        }
        catch {

        }
    }
    //#endregion

    destroySubscriptions() {
        if (this.retrieveDocumentLocationsSub)
            this.retrieveDocumentLocationsSub.unsubscribe();
        if (this.getFilesSub)
            this.getFilesSub.unsubscribe();
        if (this.createFoldersSub)
            this.createFoldersSub.unsubscribe();
        if (this.uploadFileSub)
            this.uploadFileSub.unsubscribe();
        if (this.downloadSub)
            this.downloadSub.unsubscribe();
        if (this.renameFileSub)
            this.renameFileSub.unsubscribe();
        if (this.renameFolderSub)
            this.renameFolderSub.unsubscribe();
        if (this.shareLinkSub)
            this.shareLinkSub.unsubscribe();
        if (this.searchSub)
            this.searchSub.unsubscribe();
        if (this.actionOutputSub)
            this.actionOutputSub.unsubscribe();
        if (this.createFolderAndUploadFilesSub)
            this.createFolderAndUploadFilesSub.unsubscribe();
        if (this.uploadFilesStartSub)
            this.uploadFilesStartSub.unsubscribe();
        if (this.uploadFileSPSub)
            this.uploadFileSPSub.unsubscribe();
        if (this.downloadSPSub)
            this.downloadSPSub.unsubscribe();
        if (this.getFilesAngularSub)
            this.getFilesAngularSub.unsubscribe();
    }
}
