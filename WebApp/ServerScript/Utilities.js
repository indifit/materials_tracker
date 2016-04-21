var jw;
(function (jw) {
    (function (MaterialsTracker) {
        (function (Utilities) {
            var RangeUtilties = (function () {
                function RangeUtilties(r) {
                    var _this = this;
                    this.range = null;
                    this.rangeValues = null;
                    this.convertToObjectArray = function () {
                        var ret = [];

                        //Read the first row of the range as the properties for the object
                        var propertyNames = [];

                        var rangeValues;

                        if (_this.rangeValues != null) {
                            rangeValues = _this.rangeValues;
                        } else {
                            rangeValues = _this.range.getValues();
                        }

                        var firstRow = rangeValues[0];

                        for (var i = 0; i < firstRow.length; i++) {
                            if (firstRow[i].toString() !== '') {
                                var thisPropertyName = RangeUtilties.toCamelCase(firstRow[i].toString());
                                propertyNames.push(thisPropertyName);
                            }
                        }

                        for (var j = 1; j < rangeValues.length; j++) {
                            var thisObject = {};

                            for (var k = 0; k < propertyNames.length; k++) {
                                thisObject[propertyNames[k]] = rangeValues[j][k].toString();
                            }

                            ret.push(thisObject);
                        }

                        return ret;
                    };
                    if (typeof r.activate == 'function') {
                        this.range = r;
                    } else {
                        this.rangeValues = r;
                    }
                }
                RangeUtilties.toCamelCase = function (input) {
                    //Split the input at the spaces
                    var words = input.split(' ');

                    for (var i = 0; i < words.length; i++) {
                        if (i === 0) {
                            //convert the first word to lower case
                            words[i] = words[i].toLowerCase();
                        } else {
                            //Capitalise the first letter of the remaining words
                            words[i] = words[i].substr(0, 1).toUpperCase() + words[i].substring(1).toLowerCase();
                        }
                    }

                    //Re-join the array
                    return words.join('');
                };

                RangeUtilties.findRowsMatchingKey = function (rangeValues, lookupVal, keyColIndex, headerRow) {
                    if (typeof keyColIndex === "undefined") { keyColIndex = 0; }
                    var rowVals;

                    var ret = new Array();

                    if (typeof headerRow != 'undefined') {
                        ret.push(headerRow);
                    }

                    for (var i = 0; i < rangeValues.length; i++) {
                        rowVals = rangeValues[i];
                        var keyColVal = rowVals[keyColIndex];

                        if (typeof keyColVal != "undefined" && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase()) {
                            ret.push(rowVals);
                        }
                    }

                    if (ret.length > 0) {
                        return ret;
                    }

                    return null;
                };

                RangeUtilties.findFirstRowMatchingKey = function (range, lookupVal, keyColIndex) {
                    if (typeof keyColIndex === "undefined") { keyColIndex = 0; }
                    var vals = range.getValues();

                    var rowVals;

                    for (var i = 0; i < vals.length; i++) {
                        rowVals = vals[i];
                        var keyColVal = rowVals[keyColIndex];

                        if (typeof keyColVal != 'undefined' && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase()) {
                            return rowVals;
                        }
                    }
                    return null;
                };
                return RangeUtilties;
            })();
            Utilities.RangeUtilties = RangeUtilties;

            var PageSelector = (function () {
                function PageSelector(args) {
                    var _this = this;
                    this.lookupProjectFromHash = function () {
                        var hashLookupSsid = '1lMwQsbonNGycRoA-Zelm151YU6S_fididqPyTXDJmVQ';

                        //Open the spreadsheet using the ssid
                        var hashLookupSs = SpreadsheetApp.openById(hashLookupSsid);

                        var sheet = hashLookupSs.getSheetByName('MaterialsTrackers');

                        var range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());

                        var projHashRow = RangeUtilties.findFirstRowMatchingKey(range, _this.projectHash, 6);

                        if (projHashRow != null) {
                            var response = {
                                projectNumber: projHashRow[2].toString(),
                                urlHash: projHashRow[6].toString(),
                                projectName: projHashRow[1].toString(),
                                projectSsid: typeof projHashRow[5] !== 'undefined' ? projHashRow[5].toString() : ''
                            };

                            return response;
                        }

                        return null;
                    };
                    this.getPage = function () {
                        if (typeof _this.projectHash == 'undefined') {
                            return {
                                templateName: 'InvalidProjectPage',
                                data: {}
                            };
                        }

                        var projectLookupResponse = _this.lookupProjectFromHash();

                        if (projectLookupResponse == null || (typeof projectLookupResponse.projectSsid === 'undefined' || projectLookupResponse.projectSsid === '')) {
                            return {
                                templateName: 'InvalidProjectPage',
                                data: {}
                            };
                        }

                        if (typeof _this.pageHash == 'undefined') {
                            return {
                                templateName: 'InvalidPageHashPage',
                                data: {}
                            };
                        }

                        //Lookup the destination page and get any necessary data
                        var templateName = MaterialsTracker.Config.ConfigurationManager.getSetting('PageHash' + _this.pageHash);

                        if (templateName == null) {
                            return {
                                templateName: 'InvalidPageHashPage',
                                data: {}
                            };
                        }

                        var getDataMethodName = MaterialsTracker.Config.ConfigurationManager.getSetting(templateName + 'DataMethod');

                        var data = _this[getDataMethodName](projectLookupResponse);

                        return {
                            templateName: templateName,
                            data: data
                        };
                    };
                    this.getClpData = function (projectData) {
                        var data = {
                            projectData: projectData,
                            coreListData: null,
                            trades: null
                        };

                        var dataFetcher = new DataFetcher();

                        var filteredCoreListData = dataFetcher.getFilteredCoreListItems(null);

                        var trades = [];

                        var savedCoreListDataObjectArray = null;

                        if (projectData.projectSsid !== '') {
                            savedCoreListDataObjectArray = dataFetcher.getSavedCoreListItems(projectData.projectSsid);
                        }

                        //Get the trades from the core list and determine which items are saved
                        filteredCoreListData.forEach(function (coreListItem) {
                            if (trades.indexOf(coreListItem.trade.toString().trim()) === -1) {
                                trades.push(coreListItem.trade.toString().trim());
                            }

                            if (projectData.projectSsid !== '') {
                                var savedItemsOfThisType = savedCoreListDataObjectArray.filter(function (savedItem) {
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
                    };
                    if (typeof args.parameters != 'undefined') {
                        this.pageHash = args.parameter['pageHash'];
                        this.projectHash = args.parameter['projectHash'];
                    }

                    if (typeof args.projectHash != 'undefined') {
                        this.pageHash = args.pageHash;

                        this.projectHash = args.projectHash;
                    }
                }
                return PageSelector;
            })();
            Utilities.PageSelector = PageSelector;

            var DataFetcher = (function () {
                function DataFetcher() {
                    var _this = this;
                    this.getCoreListItems = function () {
                        var lastRow = DataFetcher.coreListSheet.getLastRow();

                        var lastColumn = DataFetcher.coreListSheet.getLastColumn();

                        var coreListRange = DataFetcher.coreListSheet.getRange(1, 1, lastRow, lastColumn);

                        return coreListRange.getValues();
                    };
                    this.getFilteredCoreListItems = function (filter) {
                        var coreListItems = _this.getCoreListItems();

                        var headerRow = coreListItems[0];

                        var filteredItems = coreListItems;

                        if (typeof filter != 'undefined' && filter != null) {
                            if (typeof filter.trade != 'undefined') {
                                filteredItems = RangeUtilties.findRowsMatchingKey(filteredItems, filter.trade, 0, headerRow);
                            }

                            if (typeof filter.category != 'undefined') {
                                filteredItems = RangeUtilties.findRowsMatchingKey(filteredItems, filter.category, 3, headerRow);
                            }

                            if (typeof filter.type != 'undefined') {
                                filteredItems = RangeUtilties.findRowsMatchingKey(filteredItems, filter.type, 4, headerRow);
                            }
                        }

                        //Retrieve the valid project dimensions for the filtered items
                        var allpdcsSheet = DataFetcher.centralPurchasingSs.getSheetByName('PDCs');

                        var allpdcsRange = allpdcsSheet.getRange(2, 1, allpdcsSheet.getLastRow(), allpdcsSheet.getLastColumn());

                        var rangeUtils = new RangeUtilties(filteredItems);

                        var filteredItemsObjectArray = rangeUtils.convertToObjectArray();

                        filteredItemsObjectArray.forEach(function (value) {
                            var pdcString = value.pdc;

                            //Get the pdc codes for this item
                            var pdcArray = pdcString.split(';');

                            value.pdcs = [];

                            pdcArray.forEach(function (pdcCode) {
                                var row = RangeUtilties.findFirstRowMatchingKey(allpdcsRange, pdcCode);

                                if (row != null) {
                                    value.pdcs.push({ code: pdcCode, description: row[1].toString() });
                                }
                            });
                        });

                        return filteredItemsObjectArray;
                    };
                    this.getSavedCoreListItems = function (projectSsid) {
                        var projectSs = SpreadsheetApp.openById(projectSsid);

                        var materialsTrackingSheet = projectSs.getSheetByName('Materials Tracking');

                        var savedItemsRange = materialsTrackingSheet.getRange(2, 1, materialsTrackingSheet.getLastRow(), materialsTrackingSheet.getLastColumn());

                        var savedItemsValues = savedItemsRange.getValues();

                        var rangeUtils = new RangeUtilties(savedItemsValues);

                        var savedItems = rangeUtils.convertToObjectArray();

                        var ret = [];

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
                DataFetcher.centralPurchasingSsid = MaterialsTracker.Config.ConfigurationManager.getSetting('CentralPurchasingSSID');

                DataFetcher.centralPurchasingSs = SpreadsheetApp.openById(DataFetcher.centralPurchasingSsid);

                DataFetcher.coreListSheet = DataFetcher.centralPurchasingSs.getSheetByName('CoreList');
                return DataFetcher;
            })();
            Utilities.DataFetcher = DataFetcher;

            var DataSaver = (function () {
                function DataSaver() {
                }
                DataSaver.getMaterialsTrackerSs = function (projectDetails) {
                    var pageSelector = new PageSelector(projectDetails);

                    var project = pageSelector.lookupProjectFromHash();

                    var materialsTrackerSs;

                    if (project.projectSsid !== '') {
                        //There is a project materials tracker spreasheet already - no need to create one
                        materialsTrackerSs = SpreadsheetApp.openById(project.projectSsid);
                    } else {
                        //Need to clone a copy of the materials tracker
                        var folderId = MaterialsTracker.Config.ConfigurationManager.getSetting('FolderID');

                        var folder = DriveApp.getFolderById(folderId);

                        var masterFiles = folder.getFilesByName('Materials Tracker Master');

                        var masterFile;

                        materialsTrackerSs = null;

                        if (masterFiles.hasNext()) {
                            masterFile = masterFiles.next();

                            //Get the Id of the new Materials Tracker and save it in the project hash lookup spreadsheet
                            var newFile = masterFile.makeCopy(project.projectName + ' Materials Tracker');

                            //Get the ProjectHash Lookup Spreadsheet ID
                            var hashLookupSsid = MaterialsTracker.Config.ConfigurationManager.getSetting(MaterialsTracker.Config.ConfigurationManager.projectNumberLookupSsidKey);

                            //Open the Hash Lookup spreadsheet using the ssid retrieved
                            var hashLookupSs = SpreadsheetApp.openById(hashLookupSsid);

                            var sheet = hashLookupSs.getSheets()[0];

                            var range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());

                            for (var i = 1; i <= range.getNumRows(); i++) {
                                if (range.getCell(i, 1).getValue() === project.urlHash) {
                                    //Set the value of the newly created materials tracker Spreadsheet
                                    range.getCell(i, 4).setValue(newFile.getId());
                                }
                            }

                            //Get the spreadsheet
                            materialsTrackerSs = SpreadsheetApp.openById(newFile.getId());

                            //Set the Project Name and Project Number in the appropriate cells
                            var projectDetailsSheet = materialsTrackerSs.getSheetByName('Project Details');

                            var projectNameCell = projectDetailsSheet.getRange('D3');

                            projectNameCell.setValue(project.projectName);

                            var projectNumberCell = projectDetailsSheet.getRange('D5');

                            projectNumberCell.setValue(project.projectNumber);
                        }
                    }

                    return materialsTrackerSs;
                };

                DataSaver.saveBasketData = function (projectDetails, basketItems) {
                    var materialsTrackerSs = DataSaver.getMaterialsTrackerSs(projectDetails);

                    var materialsTrackingSheet = materialsTrackerSs.getSheetByName('Materials Tracking');

                    var materialsTrackingRange = materialsTrackingSheet.getRange(1, 1, materialsTrackingSheet.getLastRow(), materialsTrackingSheet.getLastColumn());

                    var firstEmptyRowNumber = 4;

                    for (var i = firstEmptyRowNumber; i <= materialsTrackingRange.getNumRows(); i++) {
                        if (materialsTrackingRange.getCell(i, 2).getValue().toString().trim() === '') {
                            firstEmptyRowNumber = i;
                            break;
                        }
                    }

                    var headerMappings = [];

                    headerMappings.push({ basketPropName: 'itemCode', materialsTrackerColNum: 3 });
                    headerMappings.push({ basketPropName: 'itemDescription', materialsTrackerColNum: 2 });
                    headerMappings.push({ basketPropName: 'expectedPurchasePrice', materialsTrackerColNum: 12 });
                    headerMappings.push({ basketPropName: 'purchaseUom', materialsTrackerColNum: 9 });
                    headerMappings.push({ basketPropName: 'factor', materialsTrackerColNum: 10 });
                    headerMappings.push({ basketPropName: 'baseUom', materialsTrackerColNum: 11 });
                    headerMappings.push({ basketPropName: 'leadTime', materialsTrackerColNum: 23 });
                    headerMappings.push({ basketPropName: 'usage', materialsTrackerColNum: 6 });
                    headerMappings.push({ basketPropName: 'quantity', materialsTrackerColNum: 8 });

                    for (var j = 0; j < basketItems.length; j++) {
                        var basketItem = basketItems[j];
                        for (var prop in basketItem) {
                            if (basketItem.hasOwnProperty(prop)) {
                                var matchingHeaders = headerMappings.filter(function (value, index, array) {
                                    return value.basketPropName === prop;
                                });

                                if (matchingHeaders.length > 0) {
                                    //Set the value of the appropriate cell
                                    materialsTrackingRange.getCell(firstEmptyRowNumber, matchingHeaders[0].materialsTrackerColNum).setValue(basketItem[prop]);
                                } else {
                                    if (prop === 'pdc') {
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
                };
                return DataSaver;
            })();
            Utilities.DataSaver = DataSaver;
        })(MaterialsTracker.Utilities || (MaterialsTracker.Utilities = {}));
        var Utilities = MaterialsTracker.Utilities;
    })(jw.MaterialsTracker || (jw.MaterialsTracker = {}));
    var MaterialsTracker = jw.MaterialsTracker;
})(jw || (jw = {}));
//# sourceMappingURL=Utilities.js.map
