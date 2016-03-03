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
                                var regExp = new RegExp('\\s', 'g');
                                var thisPropertyName = firstRow[i].toString().replace(regExp, '');
                                thisPropertyName = thisPropertyName.substr(0, 1).toLowerCase() + thisPropertyName.substring(1);
                                propertyNames.push(thisPropertyName);
                            }
                        }

                        for (var j = 1; j < rangeValues.length; j++) {
                            var thisObject = new Object();

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
                RangeUtilties.findRowsMatchingKey = function (range, lookupVal, keyColIndex, keepHeaderRow) {
                    if (typeof keyColIndex === "undefined") { keyColIndex = 0; }
                    if (typeof keepHeaderRow === "undefined") { keepHeaderRow = false; }
                    var vals = range.getValues();
                    var rowVals = null;

                    var ret = new Array();

                    if (keepHeaderRow) {
                        ret.push(vals[0]);
                    }

                    for (var i = 0; i < vals.length; i++) {
                        rowVals = vals[i];
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
                    var rowVals = null;
                    for (var i = 0; i < vals.length; i++) {
                        rowVals = vals[i];
                        var keyColVal = rowVals[keyColIndex];

                        if (typeof keyColVal != "undefined" && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase()) {
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
                        var hashLookupSsid = jw.MaterialsTracker.Config.ConfigurationManager.getSetting(jw.MaterialsTracker.Config.ConfigurationManager.projectNumberLookupSsidKey);

                        //Open the spreadsheet using the ssid
                        var hashLookupSs = SpreadsheetApp.openById(hashLookupSsid);

                        var sheet = hashLookupSs.getSheets()[0];

                        var range = sheet.getRange(2, 1, 100, 4);

                        var projHashRow = RangeUtilties.findFirstRowMatchingKey(range, _this.projectHash);

                        if (projHashRow != null) {
                            var response = {
                                projectNumber: parseInt(projHashRow[1].toString()),
                                urlHash: projHashRow[0].toString(),
                                projectName: projHashRow[2].toString(),
                                kingdomHallAddress: projHashRow[3].toString()
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
                        var templateName = jw.MaterialsTracker.Config.ConfigurationManager.getSetting('PageHash' + _this.pageHash);

                        if (templateName == null) {
                            return {
                                templateName: 'InvalidPageHashPage',
                                data: {}
                            };
                        }

                        var getDataMethodName = MaterialsTracker.Config.ConfigurationManager.getSetting(templateName + 'DataMethod');

                        var data = _this[getDataMethodName](projectLookupResponse);

                        data.projectHash = _this.projectHash;

                        return {
                            templateName: templateName,
                            data: data
                        };
                    };
                    this.getIndexPageData = function (projectData) {
                        var data = {};

                        data['projectData'] = projectData;

                        return data;
                    };
                    this.pageHash = request.parameter['pageHash'];
                    this.projectHash = request.parameter['projectHash'];
                }
                return PageSelector;
            })();
            Utilities.PageSelector = PageSelector;
        })(MaterialsTracker.Utilities || (MaterialsTracker.Utilities = {}));
        var Utilities = MaterialsTracker.Utilities;
    })(jw.MaterialsTracker || (jw.MaterialsTracker = {}));
    var MaterialsTracker = jw.MaterialsTracker;
})(jw || (jw = {}));
//# sourceMappingURL=Utilities.js.map
