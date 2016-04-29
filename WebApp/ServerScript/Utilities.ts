module jw.MaterialsTracker.Utilities
{
    export class RangeUtilties
    {
        private range: GoogleAppsScript.Spreadsheet.Range = null;

        private rangeValues: Object[][] = null;

        constructor(range: GoogleAppsScript.Spreadsheet.Range);
        constructor(rangeValues: Object[][]);        
        constructor(r: any)
        {                                   
            if (typeof r.activate == 'function')
            {
                this.range = r;
            } else
            {
                this.rangeValues = r;
            }            
        }

        static toCamelCase = (input: string): string =>
        {
            //Split the input at the spaces
            var words: string[] = input.split(' ');            
                        
            for (var i = 0; i < words.length; i++)
            {                
                if (i === 0)
                {
                    //convert the first word to lower case            
                    words[i] = words[i].toLowerCase();
                } else
                {
                    //Capitalise the first letter of the remaining words
                    words[i] = words[i].substr(0, 1).toUpperCase() + words[i].substring(1).toLowerCase();
                }                
            }

            //Re-join the array
            return words.join('');
        };
        

        convertToObjectArray = (): Object[] =>
        {
            var ret: any = []; 

            //Read the first row of the range as the properties for the object
            var propertyNames: string[] = [];

            var rangeValues: Object[][];

            if (this.rangeValues != null)
            {
                rangeValues = this.rangeValues;
            } else
            {
                rangeValues = this.range.getValues();
            }            

            var firstRow: Object[] = rangeValues[0];

            for (var i = 0; i < firstRow.length; i++)
            {
                if (firstRow[i].toString() !== '')
                {
                    var thisPropertyName: string = RangeUtilties.toCamelCase(firstRow[i].toString());                    
                    propertyNames.push(thisPropertyName);
                }
            }

            for (var j = 1; j < rangeValues.length; j++)
            {
                var thisObject: any = {};

                for (var k = 0; k < propertyNames.length; k++)
                {
                    thisObject[propertyNames[k]] = rangeValues[j][k].toString();
                }

                ret.push(thisObject);
            }
            
            return ret;
        };

        static findRowsMatchingKey = (
            rangeValues: Object[][],
            lookupVal: string,
            keyColIndex: number = 0,
            headerRow?: Object[]
            ): Object[][] =>
        {            
            var rowVals: Object[];

            var ret: Object[][] = new Array<Array<Object>>();

            if (typeof headerRow != 'undefined')
            {
                ret.push(headerRow);
            }

            for (var i = 0; i < rangeValues.length; i++) {
                rowVals = rangeValues[i];
                var keyColVal = rowVals[keyColIndex];

                if (typeof keyColVal != "undefined" && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase())
                {
                    ret.push(rowVals);
                }
            }

            if (ret.length > 0)
            {
                return ret;
            }

            return null;
        };

        static findFirstRowMatchingKey = (
            range: GoogleAppsScript.Spreadsheet.Range,
            lookupVal: string,
            keyColIndex: number = 0            
            ): Object[]=>
        {
            var vals: Object[][] = range.getValues();            

            var rowVals: Object[];            

            for (var i = 0; i < vals.length; i++) {
                rowVals = vals[i];
                var keyColVal = rowVals[keyColIndex];                

                if (typeof keyColVal != 'undefined' && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase()) {
                    return rowVals;
                }
            }
            return null;            
        };        
    }
    
    export class PageSelector
    {
        private pageHash: string;

        private projectHash: string;
        
        constructor(args: {pageHash: string; projectHash: string});
        constructor(args: GoogleAppsScript.Script.IParameters);
        constructor(args: any)
        {            
            if (typeof args.parameters != 'undefined')
            {
                this.pageHash = args.parameter['pageHash'];
                this.projectHash = args.parameter['projectHash'];    
            }

            if (typeof args.projectHash != 'undefined')
            {             
                this.pageHash = args.pageHash;

                this.projectHash = args.projectHash;
            }
        }

        lookupProjectFromHash = (): MaterialsTracker.Interfaces.IProjectHashLookupResponse => {
            var hashLookupSsid: string = '1lMwQsbonNGycRoA-Zelm151YU6S_fididqPyTXDJmVQ';

            //Open the spreadsheet using the ssid
            var hashLookupSs: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(hashLookupSsid);

            var sheet: GoogleAppsScript.Spreadsheet.Sheet = hashLookupSs.getSheetByName('MaterialsTrackers');

            var range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());

            var projHashRow: Object[] = RangeUtilties.findFirstRowMatchingKey(range, this.projectHash, 6);

            if (projHashRow != null)
            {
                var response: MaterialsTracker.Interfaces.IProjectHashLookupResponse = {
                    projectNumber: projHashRow[2].toString(),
                    urlHash: projHashRow[6].toString(),
                    projectName: projHashRow[1].toString(),
                    projectSsid: typeof projHashRow[5] !== 'undefined' ? projHashRow[5].toString() : ''                    
                };

                return response;
            }            

            return null;
        };


        getPage = (): MaterialsTracker.Interfaces.IPage =>
        {           
            if (typeof this.projectHash == 'undefined')
            {
                return {
                    templateName: 'InvalidProjectPage',
                    data: {}
                };
            }

            var projectLookupResponse: MaterialsTracker.Interfaces.IProjectHashLookupResponse = this.lookupProjectFromHash();


            if (projectLookupResponse == null || (typeof projectLookupResponse.projectSsid === 'undefined' || projectLookupResponse.projectSsid === ''))
            {
                return {
                    templateName: 'InvalidProjectPage',
                    data: {}
                };
            }

            if (typeof this.pageHash == 'undefined')
            {
                return {
                    templateName: 'InvalidPageHashPage',
                    data: {}
                };
            }
            
            //Lookup the destination page and get any necessary data
            var templateName: string = MaterialsTracker.Config.ConfigurationManager.getSetting('PageHash' + this.pageHash);

            if (templateName == null)
            {
                return {
                    templateName: 'InvalidPageHashPage',
                    data: {}
                };
            }

            var getDataMethodName: string = MaterialsTracker.Config.ConfigurationManager.getSetting(templateName + 'DataMethod');

            var data: any = this[getDataMethodName](projectLookupResponse);

            return {
                templateName: templateName,
                data: data
            };
        };
        
        getClpData = (projectData: MaterialsTracker.Interfaces.IProjectHashLookupResponse): Object =>
        {
            var data: any = {
                projectData: projectData,
                coreListData: null,
                trades: null
            };

            var dataFetcher: DataFetcher = new DataFetcher();
            
            var filteredCoreListData: Object[] = dataFetcher.getFilteredCoreListItems(null);

            var trades: string[] = [];

            var savedCoreListDataObjectArray: Object[] = null;

            if (projectData.projectSsid !== '')
            {
                savedCoreListDataObjectArray = dataFetcher.getSavedCoreListItems(projectData.projectSsid);
            }

            //Get the trades from the core list and determine which items are saved
            filteredCoreListData.forEach((coreListItem: any): void => {
                if (trades.indexOf(coreListItem.trade.toString().trim()) === -1) {
                    trades.push(coreListItem.trade.toString().trim());
                }

                if (projectData.projectSsid !== '')
                {
                    var savedItemsOfThisType: any[] = savedCoreListDataObjectArray.filter((savedItem: any): boolean => {
                        return coreListItem.itemCode === savedItem.itemCode;
                    });

                    coreListItem.isSaved = savedItemsOfThisType.length > 0;

                    coreListItem.quantity = savedItemsOfThisType.length > 0 ? savedItemsOfThisType[0].quantity : null;   
                }
            });

                                            
            data.projectData = projectData;

            data.coreListData = JSON.stringify(filteredCoreListData);                        

            data.trades = JSON.stringify(trades);

            return data;
        }
    }   

    export class DataFetcher
    {
        private static centralPurchasingSsid: string = MaterialsTracker.Config.ConfigurationManager.getSetting('CentralPurchasingSSID');

        private static centralPurchasingSs: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(DataFetcher.centralPurchasingSsid);

        private static coreListSheet = DataFetcher.centralPurchasingSs.getSheetByName('CoreList');

        public getCoreListItems = (): Object[][] =>
        {
            var lastRow: number = DataFetcher.coreListSheet.getLastRow();

            var lastColumn: number = DataFetcher.coreListSheet.getLastColumn();

            var coreListRange: GoogleAppsScript.Spreadsheet.Range = DataFetcher.coreListSheet.getRange(1, 1, lastRow, lastColumn);

            return coreListRange.getValues();
        };

        public getFilteredCoreListItems = (filter: MaterialsTracker.Interfaces.ICoreListFilter): Object[] =>
        {
            var coreListItems: Object[][] = this.getCoreListItems();

            var headerRow: Object[] = coreListItems[0];

            var filteredItems: Object[][] = coreListItems;

            if(typeof filter != 'undefined' && filter != null){
                if(typeof filter.trade != 'undefined'){
                    filteredItems = RangeUtilties.findRowsMatchingKey(filteredItems, filter.trade, 0, headerRow);
                }

                if (typeof filter.category != 'undefined')
                {
                    filteredItems = RangeUtilties.findRowsMatchingKey(filteredItems, filter.category, 3, headerRow);
                }

                if (typeof filter.type != 'undefined')
                {
                    filteredItems = RangeUtilties.findRowsMatchingKey(filteredItems, filter.type, 4, headerRow);
                }
            }

            //Retrieve the valid project dimensions for the filtered items
            var allpdcsSheet = DataFetcher.centralPurchasingSs.getSheetByName('PDCs');

            var allpdcsRange = allpdcsSheet.getRange(2, 1, allpdcsSheet.getLastRow(), allpdcsSheet.getLastColumn());

            var rangeUtils = new RangeUtilties(filteredItems);

            var filteredItemsObjectArray: Object[] = rangeUtils.convertToObjectArray();

            filteredItemsObjectArray.forEach((value: any): void => {
                var pdcString = value.pdc;

                //Get the pdc codes for this item
                var pdcArray = pdcString.split(';');

                value.pdcs = [];

                pdcArray.forEach((pdcCode: any): void => {
                    var row: Object[] = RangeUtilties.findFirstRowMatchingKey(allpdcsRange, pdcCode);

                    if (row != null) {
                        value.pdcs.push({ code: pdcCode, description: row[1].toString() });
                    }
                });
            });


            return filteredItemsObjectArray;
        };  
        
        public getSavedCoreListItems = (projectSsid: string): Object[] =>
        {            
            var projectSs: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(projectSsid);

            var materialsTrackingSheet: GoogleAppsScript.Spreadsheet.Sheet = projectSs.getSheetByName('Materials Tracking');

            var savedItemsRange: GoogleAppsScript.Spreadsheet.Range = materialsTrackingSheet.getRange(2, 1, materialsTrackingSheet.getLastRow(), materialsTrackingSheet.getLastColumn());

            var savedItemsValues = savedItemsRange.getValues();

            var rangeUtils = new RangeUtilties(savedItemsValues);

            var savedItems: any[] = rangeUtils.convertToObjectArray();

            var ret: MaterialsTracker.Interfaces.ISavedItem[] = [];

            for (var i = 0; i < savedItems.length; i++) {
                if (savedItems[i].itemCode !== '') {
                    ret.push({
                        itemCode: savedItems[i].itemCode.toString(),
                        pdc: savedItems[i].pdCode.toString(),
                        quantity: parseInt(savedItems[i].qty.toString())
                    });
                }
            }

            return ret;                
        };   
    }

    export class DataSaver
    {
        private static getMaterialsTrackerSs = (projectDetails: { pageHash: string; projectHash: string }): GoogleAppsScript.Spreadsheet.Spreadsheet =>
        {
            var pageSelector: PageSelector = new PageSelector(projectDetails);

            var project: MaterialsTracker.Interfaces.IProjectHashLookupResponse = pageSelector.lookupProjectFromHash();

            var materialsTrackerSs: GoogleAppsScript.Spreadsheet.Spreadsheet;

            if (project.projectSsid !== '')
            {
                //There is a project materials tracker spreasheet already - no need to create one
                materialsTrackerSs = SpreadsheetApp.openById(project.projectSsid);
            } else
            {
                //Need to clone a copy of the materials tracker
                var folderId: string = MaterialsTracker.Config.ConfigurationManager.getSetting('FolderID');

                var folder: GoogleAppsScript.Drive.Folder = DriveApp.getFolderById(folderId);

                var masterFiles: GoogleAppsScript.Drive.FileIterator = folder.getFilesByName('Materials Tracker Master');

                var masterFile: GoogleAppsScript.Drive.File;

                materialsTrackerSs = null;

                if (masterFiles.hasNext())
                {
                    masterFile = masterFiles.next();

                    //Get the Id of the new Materials Tracker and save it in the project hash lookup spreadsheet
                    var newFile: GoogleAppsScript.Drive.File = masterFile.makeCopy(project.projectName + ' Materials Tracker');
                    
                    //Get the ProjectHash Lookup Spreadsheet ID
                    var hashLookupSsid: string = MaterialsTracker.Config.ConfigurationManager.getSetting(MaterialsTracker.Config.ConfigurationManager.projectNumberLookupSsidKey);

                    //Open the Hash Lookup spreadsheet using the ssid retrieved
                    var hashLookupSs: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(hashLookupSsid);

                    var sheet: GoogleAppsScript.Spreadsheet.Sheet = hashLookupSs.getSheets()[0];

                    var range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());

                    //Look for the row that matches this project
                    for (var i = 1; i <= range.getNumRows(); i++)
                    {
                        if (range.getCell(i, 1).getValue() === project.urlHash)
                        {
                            //Set the value of the newly created materials tracker Spreadsheet
                            range.getCell(i, 4).setValue(newFile.getId());                                                       
                        }
                    }
                    
                    //Get the spreadsheet
                    materialsTrackerSs = SpreadsheetApp.openById(newFile.getId());

                    //Set the Project Name and Project Number in the appropriate cells
                    var projectDetailsSheet: GoogleAppsScript.Spreadsheet.Sheet = materialsTrackerSs.getSheetByName('Project Details');

                    var projectNameCell: GoogleAppsScript.Spreadsheet.Range = projectDetailsSheet.getRange('D3');                    

                    projectNameCell.setValue(project.projectName);

                    var projectNumberCell: GoogleAppsScript.Spreadsheet.Range = projectDetailsSheet.getRange('D5');

                    projectNumberCell.setValue(project.projectNumber);                    
                }                
            }

            return materialsTrackerSs;
        }

        static saveBasketData = (projectDetails: { pageHash: string; projectHash: string }, basketItems: Object[]): void =>
        {
            var materialsTrackerSs: GoogleAppsScript.Spreadsheet.Spreadsheet = DataSaver.getMaterialsTrackerSs(projectDetails);            

            var materialsTrackingSheet: GoogleAppsScript.Spreadsheet.Sheet = materialsTrackerSs.getSheetByName('WebAppMaterials');

            var materialsTrackingRange: GoogleAppsScript.Spreadsheet.Range = materialsTrackingSheet.getRange(1, 1, materialsTrackingSheet.getLastRow(), materialsTrackingSheet.getLastColumn());

            var firstEmptyRowNumber: number = 4; //In an empty tracker this will be the first line item row

            //Find the first empty row in the item description column
            for (var i = firstEmptyRowNumber; i <= materialsTrackingRange.getNumRows(); i++)
            {
                if (materialsTrackingRange.getCell(i, 2).getValue().toString().trim() === '')
                {
                    firstEmptyRowNumber = i;
                    break;
                }
            }

            var headerMappings: { basketPropName: string; materialsTrackerColNum: number }[] = [];

            headerMappings.push({ basketPropName: 'itemCode', materialsTrackerColNum: 3 });
            headerMappings.push({ basketPropName: 'itemDescription', materialsTrackerColNum: 2 });
            headerMappings.push({ basketPropName: 'expectedPurchasePrice', materialsTrackerColNum: 12 });
            headerMappings.push({ basketPropName: 'purchaseUom', materialsTrackerColNum: 9 });
            headerMappings.push({ basketPropName: 'factor', materialsTrackerColNum: 10 });
            headerMappings.push({ basketPropName: 'baseUom', materialsTrackerColNum: 11 });
            headerMappings.push({ basketPropName: 'leadTime', materialsTrackerColNum: 23 });
            headerMappings.push({ basketPropName: 'usage', materialsTrackerColNum: 6 });
            headerMappings.push({ basketPropName: 'quantity', materialsTrackerColNum: 8 });

            //Read the properties of the basketItems and insert the data into the relevant columns
            for (var j = 0; j < basketItems.length; j++)
            {
                var basketItem = basketItems[j];
                for (var prop in basketItem)
                {
                    if (basketItem.hasOwnProperty(prop))
                    {
                        var matchingHeaders: { basketPropName: string; materialsTrackerColNum: number }[] = headerMappings.filter((value: { basketPropName: string; materialsTrackerColNum: number }, index: number, array: Object[]): boolean =>
                        {
                            return value.basketPropName === prop;
                        });

                        if (matchingHeaders.length > 0)
                        {
                            //Set the value of the appropriate cell
                            if (prop === 'usage')
                            {
                                var userEmail: string = Session.getActiveUser().getEmail();
                                var val: string = basketItem[prop];

                                if (userEmail.trim() !== '')
                                {
                                    materialsTrackingRange.getCell(firstEmptyRowNumber, matchingHeaders[0].materialsTrackerColNum).setValue(val).setNote(userEmail);
                                } else
                                {
                                    materialsTrackingRange.getCell(firstEmptyRowNumber, matchingHeaders[0].materialsTrackerColNum).setValue(val);
                                }
                                
                                
                            }
                            else
                            {
                                materialsTrackingRange.getCell(firstEmptyRowNumber, matchingHeaders[0].materialsTrackerColNum).setValue(basketItem[prop]);    
                            }
                            
                        } else
                        {
                            if (prop === 'pdc')
                            {                                
                                materialsTrackingRange.getCell(firstEmptyRowNumber, 17).setValue(basketItem['pdc'].description);
                                materialsTrackingRange.getCell(firstEmptyRowNumber, 18).setValue(basketItem['pdc'].code);                                
                            }

                            materialsTrackingRange.getCell(firstEmptyRowNumber, 4).setValue('Core Item');
                            materialsTrackingRange.getCell(firstEmptyRowNumber, 19).setValue('Branch Purchasing');
                            materialsTrackingRange.getCell(firstEmptyRowNumber, 27).setValue('1 To Be Reviewed');
                        }
                    }
                }

                firstEmptyRowNumber++;
            }
        }
    }
}