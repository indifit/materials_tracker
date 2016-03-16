var jw;
(function (jw) {
    (function (MaterialsTracker) {
        (function (Utilities) {
            var RangeUtilties = (function () {
                function RangeUtilties(r) {
                    var _this = this;
                    this.range = null;
                    this.rangeValues = null;
                    this.toCamelCase = function (input) {
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
                                var thisPropertyName = _this.toCamelCase(firstRow[i].toString());
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
                function PageSelector(request) {
                    var _this = this;
                    this.lookupProjectFromHash = function () {
                        var hashLookupSsid = MaterialsTracker.Config.ConfigurationManager.getSetting(MaterialsTracker.Config.ConfigurationManager.projectNumberLookupSsidKey);

                        //Open the spreadsheet using the ssid
                        var hashLookupSs = SpreadsheetApp.openById(hashLookupSsid);

                        var sheet = hashLookupSs.getSheets()[0];

                        var range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());

                        var projHashRow = RangeUtilties.findFirstRowMatchingKey(range, _this.projectHash);

                        if (projHashRow != null) {
                            var response = {
                                projectNumber: parseInt(projHashRow[1].toString()),
                                urlHash: projHashRow[0].toString(),
                                projectName: projHashRow[2].toString(),
                                projectSsid: typeof projHashRow[3] !== 'undefined' ? projHashRow[3].toString() : '',
                                kingdomHallAddress: typeof projHashRow[4] !== 'undefined' ? projHashRow[4].toString() : ''
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

                        if (projectLookupResponse == null) {
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
                    this.getIndexPageData = function (projectData) {
                        var data = {
                            projectData: projectData,
                            coreListData: null,
                            savedCoreListData: null,
                            trades: null
                        };

                        var dataFetcher = new DataFetcher();

                        var coreListData = dataFetcher.getCoreListItems();

                        var rangeUtils = new RangeUtilties(coreListData);

                        var coreListDataObjectArray = rangeUtils.convertToObjectArray();

                        var savedCoreListDataObjectArray = dataFetcher.getSavedCoreListItems(projectData.projectSsid);

                        var trades = [];

                        //Get the trades from the core list
                        coreListDataObjectArray.forEach(function (value) {
                            if (trades.indexOf(value.trade.toString().trim()) === -1) {
                                trades.push(value.trade.toString().trim());
                            }
                        });

                        data.projectData = projectData;

                        data.coreListData = JSON.stringify(coreListDataObjectArray);

                        data.savedCoreListData = JSON.stringify(savedCoreListDataObjectArray);

                        data.trades = JSON.stringify(trades);

                        return data;
                    };
                    this.pageHash = request.parameter['pageHash'];
                    this.projectHash = request.parameter['projectHash'];
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
        })(MaterialsTracker.Utilities || (MaterialsTracker.Utilities = {}));
        var Utilities = MaterialsTracker.Utilities;
    })(jw.MaterialsTracker || (jw.MaterialsTracker = {}));
    var MaterialsTracker = jw.MaterialsTracker;
})(jw || (jw = {}));
//# sourceMappingURL=Utilities.js.map
