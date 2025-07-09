import { Component, OnInit, Input, Pipe, PipeTransform, ViewChild, HostListener, ChangeDetectorRef } from '@angular/core';
import { DatePipe } from '@angular/common';
import { Subscription, Subject } from 'rxjs';

import { Connector } from "../../shared/connector";
import { GridService } from './grid.service';
import { EntityConfiguration } from 'src/app/shared/entity-configuration';
import { A2dAppService, SafePipe } from 'src/app/a2d-app.service';
import { UtilityService } from 'src/app/shared/utility/utility.service';
import { DropboxService } from 'src/app/dropbox/dropbox.service';
import { SharepointService } from 'src/app/sharepoint/sharepoint.service';
import { SpinnerService } from 'src/app/core/spinner/spinner.service';
import { AzureService } from 'src/app/azure/azure/azure.service';
import { BreadcrumbService } from '../breadcrumb/breadcrumb.service';
import { ModalService } from 'src/app/shared/modal/modal.service';
import { DomSanitizer } from '@angular/platform-browser';
import { GridData } from 'src/app/shared/grid-definition';
import { PrimeNGConfig } from 'primeng/api';
import { MessageService } from 'primeng/api';
import { Table } from 'primeng/table';
import { element } from 'protractor';
import { Console } from 'console';
import { subscribe } from 'diagnostics_channel';


//import {CookieService} from 'ngx-cookie-service';


@Component({
    selector: 'a2d-app-grid',
    templateUrl: './grid.component.html',
    styleUrls: ['./grid.component.css'],
    providers: [DatePipe]
})

export class GridComponent implements OnInit {
    @Input() selectedConnectorTab: Connector;
    @Input() selectedEntityConfiguration: EntityConfiguration;
    @Input() connector: Connector;
    @ViewChild('dt') table: Table;

    selectedPath: any;
    id: any;
    defaultDivStyles = { 'background-color': 'white' };
    hoveredDivStyles = { 'background-color': '#D7EBF9' };
    fileType: string;
    path: any = null;
    selectedGridData: any = [];
    isExistSub: Subscription;
    getSharePointLookUpValuesSub = new Subscription();
    //getSharePointLookUpValues$= new Subject();
    uploadFilesSub: Subscription;
    isAllSelected: boolean = false;
    //cols: any[];
    clonedRowData: { [UniqueId: string]: any } = {};
    fileurl = '';
    fileInfoTooltip: any = '';
    fieNameTooltip: any = '';
    spSite: any = "";
    _SpIFrame: any = "";
    minDate: Date;
    maxDate: Date = new Date();
    //shrujan 13 feb 23 for D&D
    moveFilesSub = new Subscription();
    traverseFilesSub = new Subscription();
    absoluteURL: any;
    draggedFile: any;
    destinationFolder: any;
    addColumns: any;
    selectedRow = null;
    selectedRows: any[];
    constructor(private gridService: GridService, public sharepointService: SharepointService,
        private dropboxService: DropboxService, public utilityService: UtilityService,
        public appService: A2dAppService, private spinnerService: SpinnerService, private azureService: AzureService,
        private breadcrumbService: BreadcrumbService, private modalService: ModalService, private sanitizer: DomSanitizer, private datePipe: DatePipe) {

    }

    ngOnInit(): void {
        if (this.selectedConnectorTab.connector_type_value != this.appService.sharepoint) {
            this.gridService.FileName = 'fileName';
            this.gridService.cols = [
                { field: 'fileName', header: this.appService.labelsMultiLanguage['thname'], width: '2rem', fieldType: 'text', isReadOnlyField: false },
                { field: 'path_display', header: this.appService.labelsMultiLanguage['thpath'], width: '2rem', fieldType: 'text', isReadOnlyField: false },
                { field: 'modified_on', header: this.appService.labelsMultiLanguage['thmodified_on'], width: '2rem', fieldType: 'datetime', isReadOnlyField: false },
                { field: 'size', header: this.appService.labelsMultiLanguage['thsize'], width: '2rem', fieldType: 'number', isReadOnlyField: false },
                { field: 'fileType', header: 'fileType', width: '2rem', fieldType: 'text', isReadOnlyField: false },
                { field: 'isChecked', header: 'isChecked', width: '2rem', fieldType: 'boolean', isReadOnlyField: false },
            ];
        }

        if (this.appService.selectedEntityRecords.length > 0) {
            this.gridService.FileName = 'fileName';
        }

        if (this.appService.isValid(this.selectedConnectorTab)) {
            this.spSite = this.utilityService.getSharePointSite(this.selectedConnectorTab.absolute_url);
            let safePipObj = new SafePipe(this.sanitizer);
            this._SpIFrame = safePipObj.transform(this.spSite);
            this._SpIFrame = this.spSite;
        }
        this.gridService.resetFiltersEmitter.subscribe(() => {
            this.resetFilters();
        });
    }

    resetFilters() {
        if (this.table) {
            this.table.reset();
        }
    }

    isEditingAllowed(field?: string, isReadOnlyField?: boolean, filetype?: string) {
        if (this.appService.isValid(this.selectedEntityConfiguration[0])) {
            if (!this.selectedEntityConfiguration[0].linearMetadataEnabled) {
                return false
            }
        }
        if (isReadOnlyField == true || field == "FileLeafRef" || field == "File_x0020_Size" || field == "Modified" || filetype == 'folder') {
            return false;
        }
        return true;

    }

    // Shrujan code starts
    /**
     * Get files on drag
     * @param event 
     */
    draggingStart(event: any, selectedFile: any) {
        let functionName: string = "draggingStart";
        try {
            event.dataTransfer.setData("text/plain", event.target.id);
            //this.draggedFile = event.target.id;
            this.draggedFile = selectedFile;
            //event.scroll= true;
            event.currentTarget.scrollTop
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
         * get destination folder at drag over
         * @param event 
         */
    dragingOver(event: any, rowData: any) {
        let functionName: string = "dragingOver";
        try {
            //this.destinationFolder = event.target.id;
            this.destinationFolder = event.currentTarget.id;
            //event.preventDefault();
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    draggingLeave(event: any, rowData: any) {
        let functionName: string = "callDragLeave";
        try {
            this.appService.drag = false;
            event.preventDefault();
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    onRowEditSave(rowData: any, path: any) {
        let functionName: string = "onRowEditSave";
        let changedRowData: any;
        let selectedRowData: any = [];
        let editEnabledRow: any = [];

        try {
            // Step 1: Check if selected rows exceed 10 before proceeding
            const selectedRows = this.utilityService.gridData.filter(x => x.isChecked == true && x.fileType == "file");
            if (selectedRows.length > 10) {
                let message = "You can only edit 10 records at a time.";
                let alertMessage = { confirmButtonLabel: "Ok", text: message, title: "Edit Limit Exceeded" };
                let alertOptions = { height: 130, width: 170 };
                Xrm.Navigation.openAlertDialog(alertMessage, alertOptions);
                return;
            }

            // Step 2: Handle valid table data
            if (this.appService.isValid(this.table.filteredValue)) {
                this.gridService.table = this.table;
            }

            changedRowData = this.compareObjects(this.clonedRowData[rowData.UniqueId as string], rowData);

            // Step 4: Prepare selected row data for upload
            if (this.utilityService.gridData.length > 0) {
                this.selectedGridData = this.utilityService.gridData.filter(x => x.isChecked == true && x.fileType == "file");
                selectedRowData = this.selectedGridData;

                const isRowPresent = selectedRowData.some(item => item.ID == rowData.ID);
                if (!isRowPresent) {
                    selectedRowData.push(rowData);
                }
            }

            // Step 5: Validate again before upload
            if (selectedRowData.length > 10) {
                let message = "You can only edit 10 records at a time.";
                let alertMessage = { confirmButtonLabel: "Ok", text: message, title: "Edit Limit Exceeded" };
                let alertOptions = { height: 130, width: 170 };
                Xrm.Navigation.openAlertDialog(alertMessage, alertOptions);
                return;
            }

            // Step 6: Upload metadata if valid
            if (selectedRowData.length > 0) {
                for (let i = 0; i < selectedRowData.length; i++) {
                    this.sharepointService.uploadMetadataToSharePoint(
                        this.selectedConnectorTab,
                        this.selectedEntityConfiguration[0],
                        selectedRowData[i].ID,
                        changedRowData,
                        path,
                        this.gridService.selectedView,
                        selectedRowData[i],
                        functionName
                    );
                }
            }

            // Step 7: Clean up cloned row data and deactivate edit mode
            delete this.clonedRowData[rowData.UniqueId as string];

            editEnabledRow = this.utilityService.gridData.filter(x => x.isEditActive == true);
            editEnabledRow.forEach(row => {
                row.isEditActive = false;
            });

        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }



    onRowEditInit(rowData: any, index: any, dt: any) {
        let functionName: string = "onRowEditInit";
        let alertMessage;
        let alertOptions;
        let message: String;
        let okMessage: string = "Ok";
        let lineData: any = [];
        let self: any;
        let editEnabledRow: any = [];
        let isCurrentRecordChecked: boolean = false;

        try {
            self = this;
            editEnabledRow = this.utilityService.gridData.filter(x => x.isEditActive == true);
            isCurrentRecordChecked = rowData.isChecked;

            if (editEnabledRow.length > 0 && !rowData.isEditActive) {
                dt.cancelRowEdit(rowData);
                message = "Edit is already enabled for one of the rows.";
                alertMessage = { confirmButtonLabel: okMessage, text: message, title: "Edit" };
                alertOptions = { height: 130, width: 170 };
                Xrm.Navigation.openAlertDialog(alertMessage, alertOptions);
                rowData.isEditEnabled = true;
                return;
            }
            rowData.isEditActive = true;
            lineData = this.utilityService.gridData.filter(x => x.isChecked == true);

            if (lineData.length > 10) {
                dt.cancelRowEdit(rowData);
                message = "You cannot select more than 10 records at a time to edit.";
                alertMessage = { confirmButtonLabel: okMessage, text: message, title: "Edit" };
                alertOptions = { height: 130, width: 170 };
                Xrm.Navigation.openAlertDialog(alertMessage, alertOptions);
                rowData.isEditEnabled = true;
                return;
            }

            if (lineData.length > 0 && !isCurrentRecordChecked) {
                dt.cancelRowEdit(rowData);
                message = "You cannot edit the unselected record.";
                alertMessage = { confirmButtonLabel: okMessage, text: message, title: "Edit" };
                alertOptions = { height: 130, width: 170 };
                Xrm.Navigation.openAlertDialog(alertMessage, alertOptions);
                rowData.isEditEnabled = true;
                rowData.isEditActive = false;
                return;
            }

            if (lineData.length <= 10) {
                lineData.forEach(row => {
                    if (row.ID != rowData.ID) {
                        row.Disable = true;
                    }
                });
                this.clonedRowData[rowData.UniqueId as string] = { ...this.deepClone(rowData, 2) };
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }


    onRowEditCancel(rowData: any, index: any, dt: any) {
        let functionName: string = "onRowEditCancel";
        let editEnabledRow: any = [];
        try {
            const index = this.utilityService.gridData.findIndex(item => item.UniqueId === rowData.UniqueId);
            if (index !== -1) {
                // Update the properties of the item
                this.utilityService.gridData[index] = this.clonedRowData[rowData.UniqueId as string];
            }
            if (this.appService.isValid(dt.filteredValue)) {
                const index = dt.filteredValue.findIndex(item => item.UniqueId === rowData.UniqueId);
                if (index !== -1) {
                    // Update the properties of the item
                    dt.filteredValue[index] = this.clonedRowData[rowData.UniqueId as string];
                }
            }
            // this.utilityService.gridData[index] = this.clonedRowData[rowData.UniqueId as string];
            delete this.clonedRowData[rowData.UniqueId as string];
            editEnabledRow = this.utilityService.gridData.filter(x => x.isEditActive == true);
            editEnabledRow.forEach(row => {
                row.isEditActive = false;
            })
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    onChangeInput(rowData: any, fieldType: string) {
        if (fieldType == 'number' || fieldType == 'currency') {
            rowData.label = rowData.value;
        }
        // if (fieldType == 'datetime' || fieldType == 'dateonly') {
        //     if (this.appService.isValid(rowData.label)) {
        //         rowData.value = new Date(rowData.label);
        //     }
        // }
    }

    formatLabel(rowData: any) {
        if (this.appService.isValid(rowData)) {
            return rowData.hasOwnProperty('label') ? rowData.label : rowData;
        }
        return rowData;
    }

    sortColumns(field: string, fieldType: string) {
        if (fieldType == 'datetime' || fieldType == 'dateonly' || fieldType == 'currency' || fieldType == 'number') {
            console.log(field + '.value');
            return field + '.value';
        }
        else {
            return field;
        }
    }

    deepClone(obj: any, depth = Infinity) {
        if (depth === 0 || obj === null || typeof obj !== 'object') {
            return obj; // Return non-object types and null as is, or if depth limit reached
        }

        // Create a new object or array to hold the cloned properties
        var clone = Array.isArray(obj) ? [] : {};

        // Iterate through each property in the original object
        for (var key in obj) {
            if (obj.hasOwnProperty(key)) {
                // Recursively clone nested objects or arrays with reduced depth
                clone[key] = this.deepClone(obj[key], depth - 1);
            }
        }

        return clone;
    }

    compareObjects(obj1: any, obj2: any) {
        const differences = {};

        function deepCompare(obj1, obj2, path = '') {
            // Check if both are objects
            if (typeof obj1 === 'object' && typeof obj2 === 'object' && obj1 !== null && obj2 !== null) {
                // Get keys of both objects
                const keys1 = Object.keys(obj1);
                const keys2 = Object.keys(obj2);

                // Find keys present in obj1 but not in obj2
                keys1.forEach(key => {
                    if (!keys2.includes(key)) {
                        differences[path + key] = { obj1Value: obj1[key], obj2Value: undefined };
                    }
                });

                // Find keys present in obj2 but not in obj1
                keys2.forEach(key => {
                    if (!keys1.includes(key)) {
                        differences[path + key] = { obj1Value: undefined, obj2Value: obj2[key] };
                    }
                });

                // Check values recursively
                for (let key of keys1) {
                    deepCompare(obj1[key], obj2[key], path + key + '.');
                }
            } else {
                // If not objects, compare values including null and Date objects
                if (obj1 !== obj2) {
                    if (obj1 instanceof Date && obj2 instanceof Date) {
                        if (obj1.getTime() !== obj2.getTime()) {
                            differences[path.split(".")[0]] = obj2
                        }
                    } else if (obj1 === null || obj2 === null) {
                        if (obj1 !== obj2) {
                            differences[path.split('.')[0]] = obj2
                        }
                    } else {
                        differences[path.split('.')[0]] = obj2;
                    }
                }
            }
        }

        deepCompare(obj1, obj2);
        return differences;
    }

    calculateMonthBounds() {
        const today = new Date();
        const currentYear = today.getFullYear();
        const currentMonth = today.getMonth();
        // Setting minDate to the first day of the current month
        this.minDate = new Date(currentYear, currentMonth, 1);
        // Setting maxDate to the last day of the current month
        this.maxDate = new Date(currentYear, currentMonth + 1, 0); // 0 here sets the date to the last day of the current month
    }

    getRowAlignedClass(fieldType: string) {
        let functionName: "getRowAlignedClass";
        try {
            if (fieldType == 'text' || fieldType == 'choice' || fieldType == 'boolean' || fieldType == 'user' || fieldType == 'multiline' || fieldType == 'url' || fieldType == 'computed') {
                return "ui-resizable-column tdCustom mobileWidthCss"
            }
            else {
                return "ui-resizable-column thCustomClass tdCustom mobileWidthCss"
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    isBottomReached: boolean = true; // Flag to track if bottom is reached

    @HostListener("window:scroll", [])
    onScroll(event: any) {
        let functionName: "onScroll";
        try {
            const element = event.target; // Use documentElement for window scrolling
            //const element = event.target;
            if (!(element.offsetHeight + element.scrollTop >= element.scrollHeight)) {
                // Not at the bottom
                this.isBottomReached = false; // Reset flag
                //this.sharepointService.nextPage();
            }
            else {
                // At the bottom
                if (!this.isBottomReached) {
                    // Perform action only if not already performed after reaching bottom
                    this.isBottomReached = true; // Set flag to true
                    if (this.sharepointService.nextRowDataHref) {
                        this.selectedRows = this.utilityService.gridData.filter(row => row.isChecked === true);
                        this.utilityService.gridData.forEach(row => {
                            if (row.isChecked === true) {
                                this.selectedRows.push(row);
                            }
                        });
                        this.sharepointService.getSharePointData(this.selectedConnectorTab, this.selectedEntityConfiguration[0], this.breadcrumbService.bcVisibleList[this.breadcrumbService.bcVisibleList.length - 1].path, this.gridService.selectedView, null, "onScroll");
                    }
                }
            }
            //this.sharepointService.getSharePointData(this.selectedConnectorTab, this.selectedEntityConfiguration[0], this.breadcrumbService.bcVisibleList[this.breadcrumbService.bcVisibleList.length - 1].path, "","onScroll");

        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }



    //Shrujan 13 feb 23 for D&D
    /**
     * This function perform drag and MOVE files from grid.
     * @returns 
     */
    // callMoveFileFromGrid(): void {
    //     let functionName: string = "callMoveFileFromGrid";
    //     let workItems = [];
    //     let folderPathArray: string[] = [];
    //     this.traverseFilesSub = new Subscription();
    //     this.isExistSub = new Subscription();
    //     this.moveFilesSub = new Subscription();
    //     let draggedFile = this.draggedFile;
    //     let destinationfolder = this.destinationFolder;
    //     try {
    //         if (this.selectedConnectorTab.connector_type_value == this.appService.sharepoint && !this.appService.isSyncCompletedForSp) {
    //             event.preventDefault();
    //             return;
    //         }
    //         if (this.appService.isValid(draggedFile) && this.appService.isValid(destinationfolder)) {
    //             if (this.appService.isValid(this.selectedConnectorTab)) {
    //                 this.spinnerService.show();
    //                 switch (this.selectedConnectorTab.connector_type_value) {
    //                     case this.appService.sharepoint:
    //                         if (this.appService.isSyncCompletedForSp) {
    //                             // To move draged file into destination folder of SharePoint
    //                             this.sharepointService.MoveFilesToSp(draggedFile, destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
    //                         }
    //                         break;
    //                     case this.appService.azurestorage:
    //                         // To move draged file into destination folder of Azure Blob Storage
    //                         this.azureService.MoveFilesToAzure(draggedFile, destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
    //                         break;
    //                     case this.appService.dropbox:
    //                         // To move draged file into destination folder of dropbox
    //                         this.dropboxService.MoveFilesToDB(draggedFile, destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
    //                         break;
    //                     case this.appService.adlsgen2:
    //                         break;
    //                     default:
    //                         break;
    //                 }
    //             }
    //             else {
    //                 console.log("Invalid Connector");
    //             }
    //         }
    //     }
    //     catch (error) {
    //         this.utilityService.throwError(error, functionName);
    //     }
    // }
    //Shrujan 13 feb 23 for D&D
    /**
     * This function perform drag and MOVE files from grid.
     * @returns 
     */
    callMoveFileFromGrid(): void {
        let functionName: string = "callMoveFileFromGrid";
        let workItems = [];
        let folderPathArray: string[] = [];
        this.traverseFilesSub = new Subscription();
        this.isExistSub = new Subscription();
        this.moveFilesSub = new Subscription();
        let draggedFile = this.draggedFile;
        let destinationfolder = this.destinationFolder;
        let data: any = {};
        try {
            if (this.selectedConnectorTab.connector_type_value == this.appService.sharepoint && !this.appService.isSyncCompletedForSp) {
                event.preventDefault();
                return;
            }
            this.selectedGridData = this.utilityService.gridData.filter(x => x.isChecked == true);
            //shrujan
            if (this.selectedGridData.length > 0 && this.selectedConnectorTab.connector_type_value != this.appService.dropbox) {
                this.spinnerService.show();
                for (let k = 0; k < this.selectedGridData.length; k++) {
                    data = this.getFilePathWithSelectedGridData();
                    switch (this.selectedConnectorTab.connector_type_value) {
                        case this.appService.sharepoint:
                            if (data["selectedGridData"][k]["fileType"] == "file")
                                this.sharepointService.MoveFilesToSp(data["selectedGridData"][k], destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
                            break;
                        case this.appService.azurestorage:
                            // To move draged file into destination folder of Azure Blob Storage
                            this.azureService.MoveFilesToAzure(data["selectedGridData"][k]["path_display"], destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
                            break;
                        case this.appService.dropbox:
                            // To move draged file into destination folder of dropbox
                            this.dropboxService.MoveFilesToDB(data["selectedGridData"][k]["path_display"], destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
                            break;
                        case this.appService.adlsgen2:
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (this.appService.isValid(draggedFile) && this.appService.isValid(destinationfolder)) {
                if (this.appService.isValid(this.selectedConnectorTab)) {
                    this.spinnerService.show();
                    switch (this.selectedConnectorTab.connector_type_value) {
                        case this.appService.sharepoint:
                            if (this.appService.isSyncCompletedForSp) {
                                // To move draged file into destination folder of SharePoint
                                if (draggedFile.fileType == 'file')
                                    this.sharepointService.MoveFilesToSp(draggedFile, destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
                            }
                            break;
                        case this.appService.azurestorage:
                            // To move draged file into destination folder of Azure Blob Storage
                            this.azureService.MoveFilesToAzure(draggedFile, destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
                            break;
                        case this.appService.dropbox:
                            // To move draged file into destination folder of dropbox
                            this.dropboxService.MoveFilesToDB(draggedFile, destinationfolder, this.selectedConnectorTab, this.selectedEntityConfiguration[0]);
                            break;
                        case this.appService.adlsgen2:
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    //Shrujan code ends

    /**
     * Get file class depend on extension
     * @param fileName 
     */
    getFileExtensionClass(file: any) {
        let functionName: "getFileExtensionClass";
        let className: any = "file empty";
        let extension: any = "";
        try {

            if (file.fileType == 'file') {
                extension = (this.appService.selectedEntityRecords.length == 0) ? file[this.gridService.FileName].split('.').pop() : file.fileName.split('.').pop();
            }
            else {
                extension = 'folder';
            }
            className = "file " + extension + "";
        } catch (error) {
            this.utilityService.throwError(error, functionName);

        }
        return className;

    }

    //Shrujan 1-Nov22 for "Preview".
    /**
     * Handling logic for previews and also for unsupported preview files
     * @param fileName 
     * @param data 
     * @returns 
     */
    getFilePreview(fileData: any): any {
        let functionName = "getFilePreview  :-";
        let extension: any = "";
        let fileSize: any;
        let modefiedOn: string = '';
        let supportedFiles: any[] = ["docx", "jpeg", "jpg", "png", "pdf", "pptx"];
        let videoTypeFiles: any[] = ["mp4", "flv", "avi", "mov", "mpg", "mkv", "wmk"];
        let imageTypeFiles: any[] = ["svg", "tiff", "bmp", "gif", "tif", "avif"]
        let unsuppHandledFile: any[] = ["xml", "doc", "xlsx", "xls", "mp3", "txt", "zip"]
        let self = this;
        let fileName: string;
        try {

            if (this.appService.isValid(this.selectedConnectorTab) && this.selectedConnectorTab.connector_type_value == 966620000) {
                if ((fileData.fileType != 'folder')) {
                    fileName = this.appService.isValid(fileData.FileLeafRef) ? fileData.FileLeafRef : fileData.fileName;
                    extension = fileName.split('.').pop();
                    if (self.selectedConnectorTab.connector_type_value == self.appService.sharepoint && self.appService.selectedEntityRecords == 0) {
                        fileSize = fileData.File_x0020_Size.label.split('.')[0];
                        modefiedOn = fileData.Modified.label;
                    }
                    else {
                        fileSize = fileData.size.split('.')[0];
                        modefiedOn = fileData.modified_on;
                    }

                    if (supportedFiles.includes(extension)) {
                        this.fileurl = "supported";
                        this.fileInfoTooltip = 'Size: ' + fileSize + ' KB, Modified On: ' + modefiedOn + '';
                        this.fieNameTooltip = fileName;
                    }
                    else if (videoTypeFiles.includes(extension)) {
                        extension = 'prevideos';
                        this.fileurl = "priviewfile " + extension + "";
                        this.fileInfoTooltip = 'Size: ' + fileSize + ' KB, Modified On: ' + modefiedOn + '';
                        this.fieNameTooltip = fileName;
                    }
                    else if (imageTypeFiles.includes(extension)) {
                        extension = 'preimages';
                        this.fileurl = "priviewfile " + extension + "";
                        this.fileInfoTooltip = 'Size: ' + fileSize + ' KB, Modified On: ' + modefiedOn + '';
                        this.fieNameTooltip = fileName;
                    }
                    else if (unsuppHandledFile.includes(extension)) {
                        this.fileurl = "priviewfile " + extension + "";
                        this.fileInfoTooltip = 'Size: ' + fileSize + ' KB, Modified On: ' + modefiedOn + '';
                        this.fieNameTooltip = fileName;
                    }
                    else {
                        extension = 'otherfiles';
                        this.fileurl = "priviewfile " + extension + "";
                        this.fileInfoTooltip = 'Size: ' + fileSize + ' KB, Modified On: ' + modefiedOn + '';
                        this.fieNameTooltip = fileName;
                    }
                }
                else {
                    // for folder 
                    fileName = this.appService.isValid(fileData.FileLeafRef) ? fileData.FileLeafRef : fileData.fileName;
                    extension = 'priviewfolder';
                    this.fileurl = "priviewfile " + extension + "";
                    this.fileInfoTooltip = fileName;
                    this.fieNameTooltip = fileName;
                }
            }
            else {
                this.utilityService.throwError("error", functionName);
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return this.fileurl;
    }

    /**
     * If previes not available for unsupoorted files then hide broken image
     * @param event 
     */
    handleMissingImage(event: Event) {
        (event.target as HTMLImageElement).style.visibility = 'hidden';
    }

    //this will hide and show views grid based on views selected option.
    showThumbnailView(selectedConnectorTab: Connector) {
        let functionName: any = "showThumbnailView";
        let isShowThumbnail: boolean = false;
        let thumbnailView: string = "Thumbnail View";
        try {
            // When Connector is sharepoint
            if (this.appService.isValid(selectedConnectorTab) && selectedConnectorTab.connector_type_value == this.appService.sharepoint) {
                // When First time loaded Views showing according to default view 
                if (!this.appService.isValid(this.gridService.viewDropdownOption)) {
                    if (selectedConnectorTab.defaultView == 966620001) {
                        isShowThumbnail = true
                    }
                    else if (selectedConnectorTab.defaultView == 966620000) {
                        isShowThumbnail = false;
                    }
                }
                else { // When user choose another option from Views dropdown.
                    if (this.gridService.viewDropdownOption == "1: 12") {
                        isShowThumbnail = true;
                    }
                    else if (this.gridService.viewDropdownOption == " Thumbnail View ") {
                        isShowThumbnail = true;
                    }
                    else {
                        isShowThumbnail = false;
                    }
                }
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return isShowThumbnail;
    }
    //

    // Shrujan done for "Preview".

    /**
     * Get delete button class depend on some conditions
     * @param path 
     * @param filetype 
     */
    getClassForDeleteButton(path: any) {
        let functionName: "getClassForDeleteButton";
        let className: any = "hideButton";
        let extension: any = "";
        try {

            // if (this.appService.securityTemplateCount > 0 && this.gridService.selectedGridData.length > 0) {
            //     if (this.gridService.selectedGridData[0].path_display == this.selectedPath && path == this.selectedPath && filetype.toLowerCase() == 'file') {
            //         className = "showButton";
            //     }
            // } else {
            //     if (this.appService.securityTemplateCount > 0 && path == this.selectedPath && filetype.toLowerCase() == 'file') {
            //         className = "showButton";
            //     }
            // }

            if (this.appService.securityTemplateCount > 0 && this.gridService.selectedGridData.length > 0) {
                if (this.gridService.selectedGridData[0].path_display == this.selectedPath && path == this.selectedPath) {
                    className = "showButton";
                }
            } else {
                if (this.appService.securityTemplateCount > 0 && path == this.selectedPath) {
                    className = "showButton";
                }
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return className;
    }

    /**
     * Disable buttons based on Security templates
     * @param selectedConnectorTab 
     * @param actionPermission 
     */
    disableButtonBasedOnSecurityTemplate(selectedConnectorTab: Connector, actionPermission: string, callingFunction?: string): boolean {
        let functionName: string = "disableButtonBasedOnSecurityTemplate";
        let disableStatus: boolean = true;
        try {
            if (this.appService.isValid(selectedConnectorTab)) {
                if (this.appService.selectedEntityRecords.length > 0 && actionPermission != 'UP') {
                    disableStatus = false;
                }
                else {
                    if (this.appService.securityTemplateCount > 0) {
                        if (this.appService.isValid(selectedConnectorTab)) {
                            if (this.appService.isValid(this.selectedConnectorTab.visible_button_list)) {
                                disableStatus = (this.selectedConnectorTab.visible_button_list.indexOf(actionPermission) > -1) ? (callingFunction == "th" || callingFunction == "td") ? false : true : (callingFunction == "th" || callingFunction == "td") ? true : false; this.selectedGridData = this.utilityService.gridData.filter(x => x.isChecked == true);
                                if (this.selectedGridData.length > 1) {
                                    disableStatus = false;
                                }
                            }
                            else {
                                if (this.selectedConnectorTab.visible_button_list == "") {
                                    disableStatus = false;
                                }
                                else if (actionPermission.toLowerCase() == 'de') { disableStatus = false; }
                            }
                        }
                    }
                }
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return disableStatus;
    }

    /**
     delete of files open modal for rename file or folder
     * @param selectedConnectorTab 
     * @param visibleList 
     * @param selectedEntityConfiguration 
     */
    delete(selectedConnectorTab: Connector, visibleList: any, selectedEntityConfiguration: EntityConfiguration, event: any): void {
        let functionName: string = "delete";
        let newName: any;
        let renameFilePath: any;
        let data: any = {};
        let self: any = this;
        try {
            if (this.appService.isValid(selectedConnectorTab)) {
                this.modalService.openConfirmWarningDialog("", (onOKClick) => {
                    if (event.isChecked == false) {
                        event.isChecked = !event.isChecked;
                    }
                    if (this.appService.currentEntityName.toLowerCase() != 'email' || this.appService.allowActivityFolderCreation) {
                        this.appService.search = false;
                        this.appService.searchWord = null;
                        this.utilityService.folderExist = false;
                        this.utilityService.fileExist = false;
                        this.selectedGridData = this.utilityService.gridData.filter(x => x.isChecked == true);
                        //Get file name path of selected row from grid
                        if (this.selectedGridData.length > 0) {
                            this.spinnerService.show();
                            for (let k = 0; k < this.selectedGridData.length; k++) {
                                data = this.getFilePathWithSelectedGridData();
                                switch (selectedConnectorTab.connector_type_value) {
                                    case this.appService.dropbox:
                                        (function (k) {
                                            setTimeout(function () {
                                                self.dropboxService.deleteIfExist(selectedConnectorTab, data["renameFilePath"], data["selectedGridData"][k]["fileName"], selectedEntityConfiguration, data["selectedGridData"], k, self.selectedGridData.length - 1);
                                            }, 1000 * k);
                                        })(k);
                                        break;
                                    case this.appService.sharepoint:
                                        //if (data["selectedGridData"][k]["fileType"] == "file")
                                        this.sharepointService.deleteFile(data["renameFilePath"], data["selectedGridData"][k]["FileLeafRef"], data["selectedGridData"][k]["FileLeafRef"], selectedConnectorTab, selectedEntityConfiguration, data["selectedGridData"][k]["UniqueId"], false);
                                        break;

                                    case this.appService.azurestorage:
                                        //if (data["selectedGridData"][k]["fileType"] == "file")
                                        this.azureService.delete(data["renameFilePath"], data["selectedGridData"][k]["fileName"], data["selectedGridData"][k]["fileName"], selectedConnectorTab, selectedEntityConfiguration);
                                        break;
                                    case this.appService.adlsgen2:
                                        break;
                                    default:
                                        console.log("No Valid Connector Found");
                                        break;
                                }
                            }
                        }
                        else {
                            this.spinnerService.hide();
                        }
                    }
                },
                    (error) => {
                    });
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * This function perform disable button when folder is selected to delete
     * @param selectedConnectorTab 
     */
    disableDeleteFolderButton(selectedConnectorTab: Connector): boolean {
        let functionName: string = "disableDeleteFolderButton";
        let disableStatus: boolean = false;
        try {
            if (this.appService.isValid(selectedConnectorTab)) {
                if (this.utilityService.gridData != null && this.utilityService.gridData.length != 0) {
                    this.selectedGridData = this.utilityService.gridData.filter(x => x.isChecked == true);
                    if (this.selectedGridData.length == 0) {
                        disableStatus = true;
                    }
                    else {
                        if (this.selectedGridData[0]["fileType"] != "file") {
                            disableStatus = true;
                        }
                    }
                }
                else {
                    disableStatus = true;
                }
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
        return disableStatus;
    }
    /**
     * Get file name path of selected row from grids
     */
    getFilePathWithSelectedGridData(): any {
        let fiunctionName: string = "getFilePathWithSelectedGridData";
        let data: any = {};
        let referencePath: any;
        let renameFilePath: any;
        try {
            this.selectedGridData = this.utilityService.gridData.filter(x => x.isChecked == true);
            referencePath = this.selectedGridData[0]["path_display"].split('/');
            renameFilePath = referencePath.splice(0, referencePath.length - 1).join('/');
            this.utilityService.fileExist = false;
            this.utilityService.folderExist = false;
            data["selectedGridData"] = this.selectedGridData;
            data["renameFilePath"] = renameFilePath;
        } catch (error) {
            this.utilityService.throwError(error, fiunctionName);
        }
        return data;
    }

    /**
     * To check and uncheck all files
     * @param selectedAllData 
     */
    onChangeAll(selectedAllData: any) {
        let functionName: string = "onChangeAll";
        try {
            if (this.isAllSelected == false) {
                for (let i = 0; i < selectedAllData.length; i++) {
                    selectedAllData[i]["isChecked"] = true;
                }
                this.isAllSelected = true;
            }
            else {
                for (let i = 0; i < selectedAllData.length; i++) {
                    selectedAllData[i]["isChecked"] = false;
                }
                this.isAllSelected = false;
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * This will call on change of check box
     * @param data 
     */
    onChange(data: any) {
        let functionName: string = "onChange";
        try {

            data.isChecked = !data.isChecked;

        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * This will sort the grid data
     * @param event 
     */
    customSort(event: any) {
        event.data.sort((data1, data2) => {
            let value1 = data1[event.field].hasOwnProperty('label') ? data1[event.field].label : data1[event.field];
            let value2 = data2[event.field].hasOwnProperty('label') ? data2[event.field].label : data2[event.field];
            let result = null;
            if (event.field == "size") {
                value1 = (value1 == "") ? 0 : parseFloat(value1);
                value2 = (value2 == "") ? 0 : parseFloat(value2);
            }
            if (value1 == null && value2 != null)
                result = -1;
            else if (value1 != null && value2 == null)
                result = 1;
            else if (value1 == null && value2 == null)
                result = 0;
            else if (typeof value1 === 'string' && typeof value2 === 'string')
                result = value1.localeCompare(value2);
            else
                result = (value1 < value2) ? -1 : (value1 > value2) ? 1 : 0;
            return (event.order * result);
        });
    }
    /**
     * On row select this will check the checkbox
     * @param event 
     */
    onRowSelect(event: any): void {
        if (event.data.isChecked == false) {
            event.data.isChecked = !event.data.isChecked;
            //this.table.initRowEdit(event.data)
        }
    }

    /**
    * On row Un select this will check the checkbox
    * @param event 
    */
    onRowUnselect(event: any): void {
        if (event.data.isChecked == true) {
            event.data.isChecked = !event.data.isChecked;
            //this.table.cancelRowEdit(event.data)
        }
    }

    /**
     * This function performs highlightRow when mouse over
     * @param data 
     */
    highlightRow(data: any): void {
        let functionName: string = "highlightRow";
        try {
            this.selectedPath = data.path_display;
            this.gridService.selectedGridData = this.utilityService.gridData.filter(x => x.isChecked == true);
            this.appService.showDeleteButton = true;
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }

    }

    /**
     * This function performs back to normal when mouse out
     */
    defaultRow(): void {
        let functionName: string = "defaultRow";
        try {
            this.selectedPath = null;
            this.appService.showDeleteButton = false;
        } catch (error) {
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
            if (this.appService.isValid(size)) {
                size = parseFloat(size);
                return size.toFixed(2);
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * Make checkbox true when click on row
     * @param data 
     */
    makeCheckboxTrue(data: any) {
        let functionname: string = "makeCheckboxTrue";
        try {
            if (data.checked == false) {
                data.isChecked = true;
                data.checked = true;
            }
            else {
                data.isChecked = false;
                data.checked = false;
            }
        } catch (error) {
            this.utilityService.throwError(error, functionname);
        }
    }

    /**
     * This function performs check which type of Connector is active on tab
     * @param selectedConnector 
     * @param path 
     * @param fileName 
     * @param selectedEntityConfiguration 
     * @param fileType 
     */
    checkConnectorType(selectedConnector: Connector, path: any, fileName: any, selectedEntityConfiguration: EntityConfiguration, fileType: string, onChange?: Boolean): void {
        let functionName: string = "checkConnectorType";
        try {
            if (this.appService.selectedEntityRecords.length == 0) {
                if (this.appService.isValid(selectedConnector)) {
                    if (this.appService.currentEntityName.toLowerCase() != 'email') {
                        if (fileType != "file") {
                            this.appService.searchWord = "";
                            this.appService.search = false;
                            this.gridService.callCheckConnectorType(selectedConnector, path, selectedEntityConfiguration, '', '', '', onChange, '', this.gridService.selectedView);
                        }
                        else {
                            let viewRight: Boolean = this.disableButtonBasedOnSecurityTemplate(selectedConnector, 'VW');
                            if (viewRight == true) {

                                switch (this.selectedConnectorTab.connector_type_value) {
                                    case this.appService.dropbox:
                                        this.spinnerService.show();
                                        this.dropboxService.downloadSelectedFileOrFolder(selectedConnector, path);
                                        break;
                                    case this.appService.sharepoint:
                                        // Get the SharePoint Sub Site if any
                                        var subSite = this.utilityService.getSharePointSubSite(selectedConnector.absolute_url);
                                        //get the cleared path
                                        path = this.utilityService.clearSubSiteFromPath(subSite, path);

                                        this.spinnerService.show();
                                        this.downloadOrViewFiles(selectedConnector, selectedEntityConfiguration, path, fileName)
                                        break;
                                    case this.appService.azurestorage:
                                    case this.appService.adlsgen2:
                                        this.downloadOrViewFilesAzure(selectedConnector, selectedEntityConfiguration, path, fileName);
                                        break;
                                }
                            }
                        }
                    }
                }
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * This function perform download or view of sharepoint file
     * @param selectedConnector 
     * @param selectedEntityConfiguration 
     * @param path 
     * @param fileName 
     */
    downloadOrViewFiles(selectedConnector: Connector, selectedEntityConfiguration: EntityConfiguration, path: any, fileName: any) {
        let isItWeb: boolean = false;
        let functionName: string = "downloadOrViewFiles";
        try {
            if (this.appService.isValid(selectedConnector)) {
                isItWeb = this.sharepointService.getDevice();
                this.sharepointService.viewFile$ = new Subject<any>();
                //we are sending false as parameter for not opening copy link box after creating sharelink
                this.sharepointService.viewFile(path, fileName, fileName, selectedConnector, selectedEntityConfiguration);
                this.sharepointService.viewFile$.subscribe(
                    (response) => {
                        this.appService._Xrm.openUrl(response, null);
                        this.spinnerService.hide();
                    }, (error) => {
                        this.utilityService.throwError(error, functionName);
                    }
                );
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * This function perform download or view of azure file
     * @param selectedConnector 
     * @param selectedEntityConfiguration 
     * @param path 
     * @param fileName 
     */
    downloadOrViewFilesAzure(selectedConnector: Connector, selectedEntityConfiguration: EntityConfiguration, path: any, fileName: any) {
        let isItWeb: boolean = false;
        let functionName: string = "downloadOrViewFilesAzure";
        let filecoll: any;
        try {
            if (this.appService.isValid(selectedConnector)) {
                isItWeb = this.sharepointService.getDevice();
                if (isItWeb == true) {
                    this.spinnerService.show();
                    this.azureService.downloadAzure(null, selectedConnector, selectedEntityConfiguration, true, path);
                }
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    /**
     * This function performs calling grid data depend on path
     * @param selectedConnector 
     * @param path 
     * @param fileName 
     * @param selectedEntityConfiguration 
     */
    callAfterSearchFile(selectedConnector: Connector, path: any, fileName: any,
        selectedEntityConfiguration: EntityConfiguration, fileType: any): void {
        let functionName: string = "callAfterSearchFile";
        try {
            if (this.appService.isValid(selectedConnector)) {
                if (fileType == "file") {
                    path = path.split(fileName);
                    path = path[0];
                } else {
                    path = path;
                }
                if (selectedConnector.connector_type_value == this.appService.azurestorage) {
                    if (fileType.toLowerCase() == "file") {
                        path = path.substring(0, path.length - 1);
                    }
                }
                this.appService.searchWord = "";
                this.appService.search = false;
                this.gridService.callCheckConnectorType(selectedConnector, path, selectedEntityConfiguration[0], "", 0, 0, "", "", this.gridService.selectedView);
            }
        } catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }

    ngOnDestroy(): void {
        this.destroySubscriptions();
    }

    destroySubscriptions(): void {
        if (this.sharepointService.fillGridSub != null)
            this.sharepointService.fillGridSub.unsubscribe();
        if (this.isExistSub != null)
            this.isExistSub.unsubscribe();
        if (this.uploadFilesSub != null)
            this.uploadFilesSub.unsubscribe();
    }
    //#region 

    dropdownItems: { label: string, value: string }[] = [];
    getSharePointLookUpValuesUsingAPI(connector: Connector, entityConfiguration: EntityConfiguration, col: any, type: any) {
        let functionName: string = "getSharePointLookUpValuesUsingAPI:";
        let self: any;
        try {
            self = this;
            self.dropdownItems = [{ label: `Loading...`, value: `` }];
            self.getSharePointLookUpValues$ = new Subject();
            self.sharepointService.getSharePointLookUpValues(connector, entityConfiguration, col, type);
            self.getSharePointLookUpValuesSub = self.sharepointService.getSharePointLookUpValues$.subscribe((response: any) => {
                self.dropdownItems = response;
            });
        }
        catch {

        }
    }
    //#endregion
    // constructor(private cdr: ChangeDetectorRef) {}

    //#region Added by Lakshman for bulk edit in metadata
    onCheckboxClick(rowData: any): any {
        let functionName: string = "onCheckboxClick: ";
        let lineData: any = [];
        let alertMessage;
        let alertOptions;
        let message: String;
        let okMessage: string = "Ok";
        let self: any;
        try {
            if (this.appService.isValid(rowData)) {
                if (rowData.isEditActive) {
                    rowData.isEditEnabled = true;
                    return;
                }

                lineData = this.utilityService.gridData.filter(x => x.isChecked == true);
                switch (lineData.length) {
                    case 0:
                        rowData.isEditEnabled = true;
                        break;
                    case 1:
                        rowData.isEditEnabled = true;
                        lineData[0].isEditEnabled = true;
                        break;
                    default:
                        let editEnabledCount: number = lineData.filter(x => x.isEditEnabled == true).length;
                        if (rowData.isChecked) {
                            if (editEnabledCount > 0) {
                                rowData.isEditEnabled = false;
                            }
                            else if (editEnabledCount == 0) {
                                rowData.isEditEnabled = true;
                            }
                        }
                        else {
                            if (editEnabledCount < 1) {
                                rowData.isEditEnabled = true;
                                lineData.sort((a, b) => a.index - b.index);
                                lineData[0].isEditEnabled = true;
                            }
                            else {
                                rowData.isEditEnabled = true;
                            }
                        }
                        break;
                }
            }
        }
        catch (error) {
            this.utilityService.throwError(error, functionName);
        }
    }
    //#endregion
}
