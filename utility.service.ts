import { Injectable } from '@angular/core';
import { Subject, Subscription } from "rxjs";

import { SpinnerService } from '../../core/spinner/spinner.service'
import { A2dAppService } from 'src/app/a2d-app.service';
import { Connector } from '../connector';
import { ModalService } from '../modal/modal.service';
import { GridData } from '../grid-definition';
import { EntityConfiguration } from '../entity-configuration';

const HIGH_START = 0xd800;
const HIGH_END = 0xdbff;
const LOW_START = 0xdc00;
const LOW_END = 0xdfff;

const mojibakeFixes = new Map<string, string>([
    ['â€œ', '-'], ['â€', '-'], ['â€˜', '-'], ['â€™', '-'],
    ['â€“', '-'], ['â€”', '-'], ['â€¦', '-'], ['Ã©', '-'],
    ['Ã¨', '-'], ['Ã', '-'], ['Â', '-'], ['�', '-'],
    ['\uFFFD', '-']
]);

@Injectable({
    providedIn: 'root'
})


export class UtilityService {
    lastReadTime: any = "";
    dateTime: any = "";
    arrayFiles = [];
    gridData = [];
    folderExist: boolean = false;
    fileExist: boolean = false;

    uploadFilesSub: Subscription;
    createFolderSub: Subscription;
    callGetFilesDragAndDropSub = new Subscription();

    textArea: any = "";
    copy: any = "";
    range: any = "";
    selection: any = "";

    callGetFilesDragAndDrop$ = new Subject<any>();
    constructor(private spinnerService: SpinnerService, private appService: A2dAppService,
        private modalService: ModalService) { }

    /**
     * Format file name
     * @param val 
     */
    formatFileName(val: string, connector_type: number): string {
        let functionName: string = "formatFileName";
        let formattedFileName: string = "";
        try {
            let valArr: string[] = this.isValid(val) ? val.split('.') : [];
            if (valArr.length > 1) {
                let extension: string = valArr.pop();
                let fileName: string = valArr.join('.');
                formattedFileName = this.formatNameWithSlash(fileName, connector_type);
                formattedFileName = `${formattedFileName}.${extension}`;
            }
            else {
                formattedFileName = this.formatNameWithSlash(val, connector_type);
            }
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return formattedFileName.trim();
    }

    /**
      * Format the name passed here
      * @param val 
      */
    formatNameWithSlash(val: string, connector_type: number): string {
        let functionName: string = "formatNameWithSlash";
        // let self:any = this;
        try {
            if (connector_type == this.appService.azurestorage || connector_type == this.appService.dropbox || connector_type == 0) {
                val = val ? val.replace(/[&\/\\#~%":*?<>{}|.]/g, '-') : val;
                //     if (val) {
                //    const unicodeSafePattern = /[^a-zA-Z0-9_\- \u200D\uFE0F\uD800-\uDFFF\u4e00-\u9fa5\u0800-\u4e00\u0600-\u06FF]/g;
                // val = val?.replace(unicodeSafePattern, '-');
                // }

            }
            else {
                val = val ? val.replace(/[\/\\"*:<>?/|]/g, '-') : val;
            }
            let isDropbox: boolean = connector_type == this.appService.dropbox ? true : false;
                val = this.removeSpecialCharacters(val, { isFile: false, isDropbox: isDropbox, replaceSlash: true });
            
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return val.trim();
    }

    /**
      * Format the name passed here
      * @param val 
      */
    formatNameWithOutSlash(val: string): string {
        let functionName: string = "formatNameWithOutSlash";
        // let self:any = this;
        try {
            val = val ? val.replace(/[&\\#~%":*?<>{}|.]/g, '-') : val;
            //val = val ? val.replace(/["*:<>?/\|]/g, '-') : val;
            //         if (val) {
            //    const unicodeSafePattern = /[^a-zA-Z0-9_\- \u200D\uFE0F\uD800-\uDFFF\u4e00-\u9fa5\u0800-\u4e00\u0600-\u06FF]/g;
            // val = val?.replace(unicodeSafePattern, '-');
            // }
            val = this.removeSpecialCharacters(val, { isFile: false, isDropbox: false, replaceSlash: false });

        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return val.trim();
    }

    /* StringUtilities.ts */

    // #region Constants

    // #endregion

    hasInvalidSurrogates(value: string): boolean {
        let functionName: string = "hasInvalidSurrogates";
        try {
            for (let i = 0; i < value.length; i++) {
                const cp = value.charCodeAt(i);
                if (cp >= HIGH_START && cp <= HIGH_END) {
                    if (i === value.length - 1 || !(value.charCodeAt(i + 1) >= LOW_START && value.charCodeAt(i + 1) <= LOW_END)) {
                        return true;
                    }
                    i++;
                } else if (cp >= LOW_START && cp <= LOW_END) {
                    return true;
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return false;
    }

    cleanInvalidSurrogates(input: string): string {
        let functionName: string = "cleanInvalidSurrogates";
        let result: string = "";
        try {
            for (let i = 0; i < input.length; i++) {
                const cp = input.charCodeAt(i);
                if (cp >= HIGH_START && cp <= HIGH_END) {
                    if (i + 1 < input.length && input.charCodeAt(i + 1) >= LOW_START && input.charCodeAt(i + 1) <= LOW_END) {
                        result += input[i] + input[i + 1];
                        i++;
                    }
                } else if (cp < LOW_START || cp > LOW_END) {
                    result += input[i];
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return result;
    }

    normalizeStyledCharacters(input: string): string {
        let functionName: string = "normalizeStyledCharacters";
        let result: string = "";
        try {
            if (!input) return input;

            for (const ch of input) {
                const cp = ch.codePointAt(0)!;
                const mapped = this.mapStylisedLetter(cp);
                result += mapped ?? ch;
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return result;
    }

    mapStylisedLetter(codePoint: number): string | null {
        // A-Z
        if (codePoint >= 0x1d400 && codePoint <= 0x1d419) return String.fromCharCode(0x41 + (codePoint - 0x1d400));
        if (codePoint >= 0x1d434 && codePoint <= 0x1d44d) return String.fromCharCode(0x41 + (codePoint - 0x1d434));
        if (codePoint >= 0x1d468 && codePoint <= 0x1d481) return String.fromCharCode(0x41 + (codePoint - 0x1d468));
        if (codePoint >= 0x1d5a0 && codePoint <= 0x1d5b9) return String.fromCharCode(0x41 + (codePoint - 0x1d5a0));
        if (codePoint >= 0x1d5d4 && codePoint <= 0x1d5ed) return String.fromCharCode(0x41 + (codePoint - 0x1d5d4));
        if (codePoint >= 0x1d608 && codePoint <= 0x1d621) return String.fromCharCode(0x41 + (codePoint - 0x1d608));

        // a-z
        if (codePoint >= 0x1d41a && codePoint <= 0x1d433) return String.fromCharCode(0x61 + (codePoint - 0x1d41a));
        if (codePoint >= 0x1d44e && codePoint <= 0x1d467) return String.fromCharCode(0x61 + (codePoint - 0x1d44e));
        if (codePoint >= 0x1d482 && codePoint <= 0x1d49b) return String.fromCharCode(0x61 + (codePoint - 0x1d482));
        if (codePoint >= 0x1d5ba && codePoint <= 0x1d5d3) return String.fromCharCode(0x61 + (codePoint - 0x1d5ba));
        if (codePoint >= 0x1d5ee && codePoint <= 0x1d607) return String.fromCharCode(0x61 + (codePoint - 0x1d5ee));
        if (codePoint >= 0x1d622 && codePoint <= 0x1d63b) return String.fromCharCode(0x61 + (codePoint - 0x1d622));

        // 0-9
        if (codePoint >= 0x1d7ce && codePoint <= 0x1d7d7) return String.fromCharCode(0x30 + (codePoint - 0x1d7ce));
        if (codePoint >= 0x1d7ec && codePoint <= 0x1d7f5) return String.fromCharCode(0x30 + (codePoint - 0x1d7ec));

        return null;
    }

    removeSpecialCharacters(value: string, opts: { isFile?: boolean; isDropbox?: boolean; replaceSlash?: boolean } = {}): string {

        var functionName = "removeSpecialCharacters";

        // #region Function-Level Variables
        var beautified = "";
        var invisiblePattern = /[\u200B-\u200C\u200E-\u200F\u202A-\u202E\u2060-\u206F]/g;

        // Build generic pattern dynamically based on replaceSlash flag
        var genericPattern: RegExp;

        if (opts.replaceSlash === false) {
            genericPattern = /[^a-zA-Z0-9 _\-\/\u200D\uFE0F\uD800-\uDFFF\u4E00-\u9FA5\u0800-\u4E00\u0600-\u06FF]/g;
        } else {
            genericPattern = /[^a-zA-Z0-9 _\-\u200D\uFE0F\uD800-\uDFFF\u4E00-\u9FA5\u0800-\u4E00\u0600-\u06FF]/g;
        }
        // #endregion

        try {
            if (!value) return "";

            if (this.hasInvalidSurrogates(value)) {
                value = this.cleanInvalidSurrogates(value);
            }

            value = this.normalizeStyledCharacters(value);

            mojibakeFixes.forEach(function (good, bad) {
                value = value.split(bad).join(good);
            });

            value = value.replace(invisiblePattern, "");
            value = value.replace(/U\+0000|0x00/g, "-").replace(/\u0000/g, "-");

            beautified = value;

            if (opts.isDropbox) {
                beautified = Array.prototype.map.call(beautified, function (ch) {
                    var code = ch.charCodeAt(0);
                    return code >= 32 && code <= 126 ? ch : "-";
                }).join("");

                beautified = beautified.replace(/[^A-Za-z0-9 _\-/]/g, "-");
                beautified = beautified.replace(/[\x00-\x1F\x7F]/g, "").trim();
            } else {
                beautified = beautified.replace(genericPattern, "-");
            }

            if (opts.isFile) {
                var lastDash = beautified.lastIndexOf("-");
                if (lastDash !== -1) {
                    beautified = beautified.slice(0, lastDash) + "." + beautified.slice(lastDash + 1);
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }

        return beautified;
    }



    /**
     * Get the SharePoint Site
     * @param val 
     */
    getSharePointSite(val: string): string {
        let functionName: string = "getSharePointSite";
        let sharePointSite: string = "";
        let sharePointSiteArr: string[] = [];
        try {
            if (this.isValid(val)) {
                sharePointSiteArr = val.split("/");
                // After the split, if length is not greater than 0, then go and return the val as an output
                if (sharePointSiteArr.length > 0) {
                    let splicedArray: string[] = sharePointSiteArr.splice(0, 3);
                    sharePointSite = `${splicedArray[0]}//${splicedArray[2]}`;
                }
                else {
                    sharePointSite = val;
                }
            }
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return sharePointSite;
    }

    /**
     * Get the SharePoint Sub-Site
     * @param val 
     */
    getSharePointSubSite(val: string): string {
        let functionName: string = "getSharePointSubSite";
        let sharePointSubSite: string = "";
        let sharePointSubSiteArr: string[] = [];
        try {
            if (this.isValid(val)) {
                sharePointSubSiteArr = val.split("/");
                // After the split, if length is not greater than 0, then go and return the val as an output
                if (sharePointSubSiteArr.length > 3) {
                    sharePointSubSiteArr.splice(0, 3);
                    sharePointSubSite = sharePointSubSiteArr.join("/");
                }
            }
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return sharePointSubSite;
    }

    /**
     * Cleare SubSite from the path
     */
    clearSubSiteFromPath(subSite: string, val: string): string {
        let functionName: string = "clearSubSiteFromPath";
        let clearedPath: string = "";
        try {
            if (this.isValid(subSite)) {
                clearedPath = val.toLowerCase().indexOf("/" + subSite.toLowerCase()) >= 0 ? val.replace(new RegExp("/" + subSite, 'gi'), '') : val;
                clearedPath = clearedPath.toLowerCase().indexOf(subSite.toLowerCase()) >= 0 ? clearedPath.replace(new RegExp(subSite, 'gi'), "") : clearedPath;
            }
            else {
                clearedPath = val;
            }
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return clearedPath;
    }

    /**
     * Generic function to validate the input value
     * @param val 
     */
    isValid(val: any): boolean {
        let functionName: string = "isValid";
        let isValid: boolean = false;
        try {
            if (val != '' && val != null && val != undefined && val != "undefined")
                isValid = true;
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return isValid;
    }

    /**
     * Craete text area for copy text
     * @param text 
     */
    createTextArea(text) {
        let functionName: "createTextArea";
        try {
            this.textArea = document.createElement('textArea');
            this.textArea.value = text;
            document.body.appendChild(this.textArea);
        }
        catch (error) {
            this.throwError(error, functionName);
        }
    }

    /**
     * Checking device is ios or not
     */
    isOS() {
        return navigator.userAgent.match(/ipad|iphone/i);
    }

    /**
     * This will copy text from textbox
     */
    selectText() {
        let functionName: "selectText";
        try {
            if (this.isOS()) {
                this.range = document.createRange();
                this.range.selectNodeContents(this.textArea);
                this.selection = window.getSelection();
                this.selection.removeAllRanges();
                this.selection.addRange(this.range);
                this.textArea.setSelectionRange(0, 999999);
            } else {
                this.textArea.select();
            }
        }
        catch (error) {
            this.throwError(error, functionName);
        }
    }

    /**
     * For copy the text from the clipboard
     */
    copyToClipboard() {
        document.execCommand('copy');
        document.body.removeChild(this.textArea);
    }

    /**
     * To copy any Text
     * @param val 
     */
    copyText(val: string) {
        let functionName: "selectText";
        try {
            this.createTextArea(val);
            this.selectText();
            this.copyToClipboard();
        }
        catch (error) {
            this.throwError(error, functionName);
        }
    }

    /**
     * Generic function to throw error
     * @param error 
     * @param functionName 
     */
    throwError(error: any, functionName: string): void {
        let err: string = `${functionName}: ${error.message || error.description}`;
        this.spinnerService.hide();
        throw new Error(err)
    }

    /**
     * This function will call after all files are stored in one collection and it has interval of 2 second 
     * If it has difference between last read time and current upload time then it will go to upload files function
     * @param uploadPath 
     * @param selectedConnectorTab 
     * @param selectedEntityConfiguration 
     * @param runningCount 
     * @param count 
     * @param value 
     */
    createFileCollectionOnDragAndDrop(uploadPath: any, selectedConnectorTab: any, selectedEntityConfiguration: any, runningCount?: any, count?: any, value?: any) {
        let functionName: string = "createFileCollectionOnDragAndDrop";
        let self = this;
        try {
            var x = setInterval(function () {

                if (self.lastReadTime != "") {
                    self.dateTime = new Date();
                    let sec: any = (self.dateTime - self.lastReadTime) / 100;

                    if (sec > 2) {
                        clearInterval(x);
                        self.uploadFileOrFolder(self.arrayFiles, uploadPath, selectedConnectorTab, selectedEntityConfiguration, runningCount, count, value);
                    }
                }
            }, 500);
        } catch (error) {
            this.throwError(error, functionName);
        }

    }

    /**
     * This function perform upload file or folder when drag and drop happend
     * @param files 
     * @param uploadPath 
     * @param selectedConnectorTab 
     * @param selectedEntityConfiguration 
     * @param runningCount 
     * @param count 
     * @param value 
     */
    uploadFileOrFolder(files: any, uploadPath: any, selectedConnectorTab: any, selectedEntityConfiguration: any, runningCount?: any, count?: any, value?: any): void {
        let functionName: string = "uploadFolder";
        let workItems = [];
        let result: any = {};
        let folderPathArray: string[] = [];
        this.uploadFilesSub = new Subscription();
        try {
            //This will create folder array to be created
            folderPathArray = this.createFolderArrayDragDrop(files);
            //This will create work items
            workItems = this.createWorkItemsDragDrop(files, selectedConnectorTab, selectedEntityConfiguration, uploadPath, value);
            result["runningCount"] = runningCount;
            result["count"] = count;
            result["folderPathArray"] = folderPathArray;
            result["workItems"] = workItems;
            result["uploadPath"] = uploadPath;
            this.callGetFilesDragAndDrop$.next(result);
        } catch (error) {
            this.throwError(error, functionName);
        }
    }

    /**
     * This will create folder array and workitesm and call upload files function
     * @param event 
     * @param selectedConnectorTab 
     * @param selectedEntityConfiguration 
     * @param uploadPath 
     * @param runningCount 
     * @param count 
     * @param value 
     */
    callGetFilesDragAndDrop(event: any, selectedConnectorTab: any, selectedEntityConfiguration: any, uploadPath: any, runningCount?: any, count?: any, value?: any) {
        let functionName: string = "callGetFilesDragAndDrop";
        try {
            this.appService.drag = false;
            this.callGetFilesDragAndDrop$ = new Subject<any>();
            this.arrayFiles = [];
            //for preventing back window
            event.preventDefault();
            if (event != null) {
                // Get the files and folders from the dragged items and loop through them to upload to the respective cloud storages
                let items = event.dataTransfer.items;
                if (this.isValid(items)) {
                    for (let index = 0; index < items.length; index++) {
                        let item = items[index].webkitGetAsEntry();
                        this.scanFiles(item, '');
                    }
                    this.createFileCollectionOnDragAndDrop(uploadPath, selectedConnectorTab, selectedEntityConfiguration, runningCount, count, value);
                }
                else {
                    //This error is come when drag and drop functionality does not support
                    this.spinnerService.hide();
                    this.modalService.openErrorDialog(this.appService.labelsMultiLanguage['dragdroperror'], (onOKClick) => {
                    });
                }
            }
        } catch (error) {
            //For edge different way is there to get files
            let items = event.dataTransfer.files;
            for (let index = 0; index < items.length; index++) {
                this.getFilesOnEdge(items[index]);
            }
            this.createFileCollectionOnDragAndDrop(uploadPath, selectedConnectorTab, selectedEntityConfiguration);
            //this.throwError(error, functionName);
        }
    }

    /**
     * Get files from edge browser
     * @param item 
     */
    getFilesOnEdge(item) {
        let functionName: string = "getFilesOnEdge";
        try {
            this.lastReadTime = new Date();
            let obj = {};
            obj["FilePath"] = item["webkitRelativePath"];
            obj["isDirectory"] = false;
            obj["isFile"] = true;
            obj["FileName"] = item["name"];
            obj["file"] = item;

            this.arrayFiles.push(obj);
        } catch (error) {
            this.throwError(error, functionName);
        }
    }

    /**
   * This function perform get files from directory
   * @param item 
   * @param container 
   */
    scanFiles(item, container) {
        let functionName: string = "scanFiles";
        let self = this;
        try {
            this.lastReadTime = new Date();
            let obj = {};
            obj["FilePath"] = item["fullPath"];
            obj["isDirectory"] = item["isDirectory"];
            obj["isFile"] = item["isFile"];
            obj["FileName"] = item["name"];
            if (item.isFile) {
                // Get file
                item.file(datafile => {
                    obj["file"] = datafile;
                });
            }
            this.arrayFiles.push(obj);
            if (item.isDirectory) {
                let directoryReader = item.createReader();

                directoryReader.readEntries(function (entries) {
                    self.lastReadTime = new Date();
                    entries.forEach(function (entry) {
                        self.lastReadTime = new Date();
                        self.scanFiles(entry, '');
                    });
                });
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
    }


    /**
     * Create folder array
     * @param files 
     */
    createFolderArrayDragDrop(files: any): any {
        let functionName: string = "createFolderArrayDragDrop";
        let fullfolderPath: any = '';
        let folderPath: any = '';
        let folderPathArray: string[] = [];
        try {
            for (let index = 0; index < files.length; index++) {
                if (files[index]["isDirectory"] == true) {
                    let folderPath: any = files[index]["FilePath"];
                    fullfolderPath = folderPath;
                    if (!folderPathArray.some((elem) => { return elem == fullfolderPath })) {
                        folderPathArray.push(this.formatNameWithOutSlash(fullfolderPath));
                    }
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return folderPathArray;
    }


    /**
     * Create folder array
     * @param files 
     */
    createFolderArrayUploadButton(files: any): any {
        let functionName: string = "createFolderArrayUploadButton";
        let fullfolderPath: any = '';
        let folderPathArray: string[] = [];
        try {
            for (let index = 0; index < files.length; index++) {
                let getFolder: any = files[index]["webkitRelativePath"].split(files[index].name);
                if (getFolder[0] != "") {
                    fullfolderPath = getFolder[0].substring(0, getFolder[0].length - 1);
                }
                else {
                    fullfolderPath = '';
                }
                if (!folderPathArray.some((elem) => { return elem == fullfolderPath })) {
                    folderPathArray.push(this.formatNameWithOutSlash(fullfolderPath));
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return folderPathArray;
    }

    /**
     * This function perform check blocked extensions added in crm for that connector
     * @param extension 
     * @param connector 
     */
    checkExtensionExist(extension: any, connector: Connector): boolean {
        let functionName: string = "checkExtensionExist";
        let blocked: boolean = false;
        let blockedExtension: any = [];
        try {
            blockedExtension = connector.blocked_extensions.split(';');
            for (let index = 0; index < blockedExtension.length; index++) {
                if (extension == blockedExtension[index].toLowerCase()) {
                    blocked = true;
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return blocked;
    }

    /**
       * This function perform check blocked extensions added in crm for that connector
       * @param extension 
       * @param connector 
       */
    checkExtensionExistInCRM(extension: any): boolean {
        let functionName: string = "checkExtensionExistInCRM";
        let blocked: boolean = false;
        let blockedExtension: any = [];
        try {
            blockedExtension = this.appService.crmBlockedExtension.split(';');
            for (let index = 0; index < blockedExtension.length; index++) {
                if (extension == blockedExtension[index].toLowerCase()) {
                    blocked = true;
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return blocked;
    }

    /**
     * This will create workitems while drag and drop
     * @param files 
     * @param selectedConnectorTab 
     * @param selectedEntityConfiguration 
     * @param uploadPath 
     * @param value 
     */
    createWorkItemsDragDrop(files: any, selectedConnectorTab: any, selectedEntityConfiguration: any, uploadPath: any, value?: any): any {
        let functionName: string = "createWorkItemsDragDrop";
        let folderPath: any = '';
        let getFolder: any = '';
        let workItems = [];
        let fileCount: any = 0;
        let fileExtension: any;
        let reason: any = '';
        let blocked: boolean = false;
        let self = this;
        try {
            ////selectedConnectorTab.max_size_UI=1572864;
            const maxBlob = selectedConnectorTab.max_size_UI * 1024;
            for (let index = 0; index < files.length; index++) {
                if (files[index]["isDirectory"] == false) {
                    fileCount = fileCount + 1;
                    let file = files[index];

                    fileExtension = file["file"]["name"].substring(file["file"]["name"].lastIndexOf(".") + 1, file["file"]["name"].length);
                    if (fileExtension != "ini") {
                        blocked = this.checkExtensionExist(fileExtension, selectedConnectorTab);
                        if (blocked == true) {
                            reason = 'This file is ignored because file is blocked in configured blocked extension.';
                        }
                        else if ((file["file"]["size"]) > (selectedConnectorTab.max_size_UI * 1024)) {
                            reason = 'This file is ignored because the file size is greater than configured file size.';
                        }
                        else if (file["file"]["size"] == 0) {
                            reason = 'This file is ignored because the file size is 0';
                        }
                        if (((file["file"]["size"]) < (selectedConnectorTab.max_size_UI * 1024)) && (file["file"]["size"] > 0) && (blocked == false)) {
                            if (file["FilePath"] != "") {
                                getFolder = file["FilePath"].toString().split(file["FileName"]);
                            }
                            if (getFolder != "") {
                                folderPath = getFolder[0].substring(0, getFolder[0].length - 1);
                            }
                            let offset = 0;
                            while (offset < (file["file"]["size"])) {
                                let chunkSize = Math.min(maxBlob, (file["file"]["size"]) - offset);
                                workItems.push({ file: file["file"], chunk: true, offset: offset, end: offset + chunkSize, size: chunkSize, path: folderPath });
                                offset += chunkSize;
                            }
                            workItems[workItems.length - 1].close = true;
                        }
                        else {
                            let ignoreFileDetails: any = [];
                            this.appService.IgnoreFileCount = this.appService.IgnoreFileCount + 1;
                            ignoreFileDetails["FileName"] = file["FileName"];
                            ignoreFileDetails["FilePath"] = file["FilePath"];
                            this.appService.IgnoreFileNames.push(ignoreFileDetails);
                            self.appService.logError(file, reason, selectedEntityConfiguration, '', file["FileName"], uploadPath, '', value);
                        }
                    }
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return workItems;
    }

    /**
   * Create collection based on the result
   * @param result this is an collectoion of files
   * return array
   */
    createCollectionOfFiles(result: any): any {
        let GridDataList: GridData[] = [];
        let functionName: string = "createCollectionOfFiles";
        let pathArray: any = [];
        try {
            for (let index = 0; index < result.length; index++) {
                let data: GridData = new GridData();
                if (result[index]['path'].length > 1) {
                    pathArray = result[index]['path'].split('/');
                    for (let z = 0; z < GridDataList.length; z++) {
                        if (GridDataList[z].fileName != pathArray[1]) {
                            if (this.appService.isValid(GridDataList[z].fileName)) {
                                data.fileName = pathArray[1];
                                data.fileType = "folder";
                            }
                        }
                    }
                    if (GridDataList.length == 0) {
                        data.fileName = pathArray[1];
                        data.fileType = "folder";
                    }
                }
                else {
                    let element: any = result[index]['file'];
                    data.fileName = element.name != null && element.name != "" ? element.name : "";
                    data.fileType = "file";
                    data.size = element.size != null && element.size != "" ? (parseInt(element.size) / 1024).toString() : "";
                    data.path_display = result[index]['path'] != null && result[index]['path'] != "" ? result[index]['path'] : "";
                    data.displayFileName = element.name != null && element.name != "" && element.name.length > 30 ? `${element.name.substring(0, 30)}...` : element.name;
                    data.displayPath = result[index]['path'] != null && result[index]['path'] != "" && result[index]['path'].length > 30 ? `${result[index]['path'].substring(0, 30)}...` : result[index]['path'];
                    if (data.fileType.toLowerCase() == "file") {
                        data.modified_on = element.lastModified != null && element.lastModified != "" ? element.lastModified : "";
                        data.modified_on = this.getDateTimeInUserTZ(this.appService.crmUserTimeZoneParameter["TimeZoneBias"], data.modified_on, this.appService.crmUserTimeZoneParameter["dateformatstring"], this.appService.crmUserTimeZoneParameter["timeformatstring"]);
                    }
                    data.isChecked = false;

                    data.fileUrl = this.getSPThumbnails(this.appService.ConnectorList[0], this.appService.EntityConfigurationList[0], element);//Shrujan
                }
                if ('fileName' in data) {
                    GridDataList.push(data);
                }
            }
        } catch (error) {
            this.throwError(error, functionName);
        }
        return GridDataList;
    }

    getSPThumbnails(connector: Connector, selectedEntityConfiguration: EntityConfiguration, fileData: any) {
        let functionName: string = "getThumbnails";
        let sharePointSite: any;
        let thumbnailURL: any = "";
        let pathUrl: any;
        let previewPath: any;
        let fileExtention: any;
        let supportedFiles: any[] = ["docx", "jpeg", "jpg", "png", "pdf", "pptx"];
        try {
            if (this.appService.selectedEntityRecords.length == 0) {
                if (this.appService.isValid(connector)) {
                    sharePointSite = this.getSharePointSite(connector.absolute_url);
                    if (this.appService.selectedEntityRecords.length > 0) {
                        fileExtention = fileData.name.split('.').pop();
                    }
                    else {
                        fileExtention = fileData.FileLeafRef.split('.').pop();
                    }

                    if (supportedFiles.includes(fileExtention)) {
                        // When UI is opning from homegrid with multiple records. 
                        if (this.appService.isValid(selectedEntityConfiguration) && !this.appService.isValid(fileData.FileRef)) {
                            pathUrl = connector.absolute_url + "/" + selectedEntityConfiguration.folder_path + "/" + fileData.FileLeafRef;
                            previewPath = pathUrl.replace(/([^:]\/)\/+/g, "$1");// Using url regex for accurate URL
                            thumbnailURL = sharePointSite + "/_layouts/15/getpreview.ashx?path=" + previewPath + "&resolution=3&force=1";
                        }
                        else { //When UI is opning from single records.
                            previewPath = sharePointSite + fileData.FileRef;
                            thumbnailURL = sharePointSite + "/_layouts/15/getpreview.ashx?path=" + previewPath + "&resolution=3&force=1";
                        }
                    }
                    else {
                        thumbnailURL = "";
                    }

                }
            }
            else {
                if (this.appService.isValid(connector)) {
                    sharePointSite = this.getSharePointSite(connector.absolute_url);
                    if (this.appService.selectedEntityRecords.length > 0) {
                        fileExtention = fileData.name.split('.').pop();
                    }
                    else {
                        fileExtention = fileData.fileName.split('.').pop();
                    }

                    if (supportedFiles.includes(fileExtention)) {
                        // When UI is opning from homegrid with multiple records. 
                        if (this.appService.isValid(selectedEntityConfiguration) && !this.appService.isValid(fileData.path_display)) {
                            pathUrl = connector.absolute_url + "/" + selectedEntityConfiguration.folder_path + "/" + fileData.name;
                            previewPath = pathUrl.replace(/([^:]\/)\/+/g, "$1");// Using url regex for accurate URL
                            thumbnailURL = sharePointSite + "/_layouts/15/getpreview.ashx?path=" + previewPath + "&resolution=3&force=1";
                        }
                        else { //When UI is opning from single records.
                            previewPath = sharePointSite + fileData.path_display;
                            thumbnailURL = sharePointSite + "/_layouts/15/getpreview.ashx?path=" + previewPath + "&resolution=3&force=1";
                        }
                    }
                    else {
                        thumbnailURL = "";
                    }

                }
            }
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return thumbnailURL;
    }
    /**
     * Convert UTC date time to Local date time
     * @param userTZBias 
     * @param utcDateTime 
     * @param dateFormat 
     * @param timeFormat 
     */
    getDateTimeInUserTZ(userTZBias: any, utcDateTime: any, dateFormat: any, timeFormat: any): any {
        let manipulateTZBias: any = "";
        let functionName: string = "getDateTimeInUserTZ";
        let convertedDateTime: any = "";
        try {
            let systemTZBias = new Date().getTimezoneOffset();

            userTZBias = userTZBias;

            if (userTZBias > 0) {
                manipulateTZBias = -userTZBias + systemTZBias;
                convertedDateTime = new Date(utcDateTime);
                convertedDateTime = new Date(convertedDateTime.getTime() + manipulateTZBias * 60000);
            }
            else if (userTZBias < 0) {
                manipulateTZBias = Math.abs(userTZBias) + systemTZBias;
                convertedDateTime = new Date(utcDateTime);
                convertedDateTime = new Date(convertedDateTime.getTime() + manipulateTZBias * 60000);
            }
            else {
                convertedDateTime = new Date(utcDateTime);
            }
            if (timeFormat.includes('tt')) {
                timeFormat = timeFormat.toLowerCase().replace(timeFormat, 'LT');
            }


            dateFormat = dateFormat.toUpperCase();
            //@ts-ignore
            convertedDateTime = this.isValid(dateFormat) ? moment(convertedDateTime).format(dateFormat + " " + timeFormat) : moment(convertedDateTime).format("M/D/YYYY" + " " + "LT");
        }
        catch (error) {
            this.throwError(error, functionName);
        }
        return convertedDateTime;
    }

    /**
     * This will create workitems while click on upload button
     * @param files 
     * @param selectedConnectorTab 
     * @param uploadPath 
     * @param selectedEntityConfiguration 
     * @param value 
     */
    createWorkItemsUploadButton(files: any, selectedConnectorTab: any, uploadPath: any, selectedEntityConfiguration: any, value?: any): any {
        let functionName: string = "createWorkItemsUploadButton";
        let newName: any = '';
        let workItems = [];
        let fileExtension: any;
        let folderPath: any = '';
        let getFolder: any = '';
        let reason: any = '';
        let self = this;
        try {
            ////selectedConnectorTab.max_size_UI=1572864;
            const maxBlob = selectedConnectorTab.max_size_UI * 1024;
            for (let index = 0; index < files.length; index++) {
                let file = files[index];
                let blob = file.slice(0, -1, file.type);
                let blocked: boolean = false;
                fileExtension = file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length);
                //Akshay
                if (fileExtension != "ini") {
                    blocked = this.checkExtensionExist(fileExtension, selectedConnectorTab);
                    if (blocked == true) {
                        reason = 'This file is ignored because file is blocked in configured blocked extension.';
                    }
                    else if (file.size > (selectedConnectorTab.max_size_UI * 1024)) {
                        reason = 'This file is ignored because the file size is greater than configured file size.';
                    }
                    else if (file.size == 0) {
                        reason = 'This file is ignored because the file size is 0.';
                    }
                    if ((file.size < (selectedConnectorTab.max_size_UI * 1024)) && (file.size > 0) && (blocked == false)) {
                        if (this.isValid(files[index]["webkitRelativePath"])) {
                            if (files[index]["webkitRelativePath"] != "") {
                                getFolder = files[index]["webkitRelativePath"].split(files[index].name);
                            }
                        }
                        if (getFolder == '') {
                            folderPath = '';
                        }
                        else {
                            folderPath = getFolder[0].substring(0, getFolder[0].length - 1);
                        }
                        let offset = 0;

                        while (offset < file.size) {
                            let chunkSize = Math.min(maxBlob, file.size - offset);
                            workItems.push({ file: file, chunk: true, offset: offset, end: offset + chunkSize, size: chunkSize, path: '/' + folderPath });
                            offset += chunkSize;
                        }
                        workItems[workItems.length - 1].close = true;
                    }
                    else {
                        let ignoreFileDetails: any = [];
                        this.appService.IgnoreFileCount = this.appService.IgnoreFileCount + 1;
                        ignoreFileDetails["FileName"] = file.name;
                        ignoreFileDetails["FilePath"] = uploadPath;
                        this.appService.IgnoreFileNames.push(ignoreFileDetails);
                        self.appService.logError(file, reason, selectedEntityConfiguration, '', file.name, uploadPath, '', value);
                    }
                }
            }
        } catch (error) {

            this.throwError(error, functionName);
        }
        return workItems;
    }

    /**
     * Converts base64 to byte array
     * @param base64 
     */
    base64ToArrayBuffer(base64: string): any {
        var functionName = "base64ToArrayBuffer";
        try {
            var binaryString = window.atob(base64);
            var len = binaryString.length;
            var bytes = new Uint8Array(len);
            for (var i = 0; i < len; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes.buffer;
        }
        catch (error) {
            this.throwError(error, functionName);
        }
    }

    /**
     * This will convert array buffer to base64
     * @param buffer 
     */
    arrayBufferToBase64(buffer: any): string {
        let binary = '';
        var functionName = "arrayBufferToBase64";
        try {
            let bytes = new Uint8Array(buffer);
            let len = bytes.byteLength;
            for (let i = 0; i < len; i++) {
                binary += String.fromCharCode(bytes[i]);
            }
            return window.btoa(binary);
        } catch (error) {
            this.throwError(error, functionName);
        }

    }
}

