﻿module jw.MaterialsTracker.Utilities
{
    export class RangeUtilties
    {
        private range: GoogleAppsScript.Spreadsheet.Range;

        constructor(range: GoogleAppsScript.Spreadsheet.Range)
        {
            this.range = range;
        }

        convertToObjectArray = (): Object[] =>
        {
            var ret: Object[] = [];

            //Read the first row of the range as the properties for the object
            var propertyNames: string[] = [];

            var rangeValues: Object[][] = this.range.getValues();

            var firstRow: Object[] = rangeValues[0];

            for (var i = 0; i < firstRow.length; i++)
            {
                if (firstRow[i].toString() !== '')
                {
                    var regExp = new RegExp('\\s', 'g');
                    var thisPropertyName: string = firstRow[i].toString().replace(regExp, '');
                    thisPropertyName = thisPropertyName.substr(0, 1).toLowerCase() + thisPropertyName.substring(1);
                    propertyNames.push(thisPropertyName);
                }
            }

            for (var j = 1; j < rangeValues.length; j++)
            {
                var thisObject = new Object();

                for (var k = 0; k < propertyNames.length; k++)
                {
                    thisObject[propertyNames[k]] = rangeValues[j][k].toString();
                }

                ret.push(thisObject);
            }
            
            return ret;
        };

        static findFirstRowMatchingKey = (
            range: GoogleAppsScript.Spreadsheet.Range,
            lookupVal: string,
            keyColIndex: number = 0            
            ): Object[]=>
        {
            var vals: Object[][] = range.getValues();
            var rowVals: Object[] = null;
            for (var i = 0; i < vals.length; i++)
            {
                rowVals = vals[i];
                var keyColVal = rowVals[keyColIndex];                

                if (typeof keyColVal != "undefined" && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase())
                {
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

        constructor(request: GoogleAppsScript.Script.IParameters)
        {
            this.pageHash = request.parameter['pageHash'];
            this.projectHash = request.parameter['projectHash'];
        }

        lookupProjectFromHash = (): jw.MaterialsTracker.Interfaces.IProjectHashLookupResponse => {
            var hashLookupSsid: string = jw.MaterialsTracker.Config.ConfigurationManager.getSetting(jw.MaterialsTracker.Config.ConfigurationManager.projectNumberLookupSsidKey);

            //Open the spreadsheet using the ssid
            var hashLookupSs: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(hashLookupSsid);

            var sheet: GoogleAppsScript.Spreadsheet.Sheet = hashLookupSs.getSheets()[0];

            var range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(2, 1, 100, 4);

            var projHashRow: Object[] = RangeUtilties.findFirstRowMatchingKey(range, this.projectHash);

            if (projHashRow != null)
            {
                var response: jw.MaterialsTracker.Interfaces.IProjectHashLookupResponse = {
                    projectNumber: parseInt(projHashRow[1].toString()),
                    urlHash: projHashRow[0].toString(),
                    projectName: projHashRow[2].toString(),
                    kingdomHallAddress: projHashRow[3].toString()
                };

                return response;
            }            

            return null;
        };


        getPage = (): jw.MaterialsTracker.Interfaces.IPage =>
        {           
            if (typeof this.projectHash == 'undefined')
            {
                return {
                    templateName: 'InvalidProjectPage',
                    data: {}
                };
            }

            var projectLookupResponse: jw.MaterialsTracker.Interfaces.IProjectHashLookupResponse = this.lookupProjectFromHash();            

            if (projectLookupResponse == null)
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
            var templateName: string = jw.MaterialsTracker.Config.ConfigurationManager.getSetting('PageHash' + this.pageHash);

            if (templateName == null)
            {
                return {
                    templateName: 'InvalidPageHashPage',
                    data: {}
                };
            }

            var getDataMethodName: string = MaterialsTracker.Config.ConfigurationManager.getSetting(templateName + 'DataMethod');

            var data: Object = this[getDataMethodName](projectLookupResponse);

            return {
                templateName: templateName,
                data: data
            };
        };
        
        getIndexPageData = (projectData: jw.MaterialsTracker.Interfaces.IProjectHashLookupResponse): Object =>
        {
            var data = new Object();

            data['projectData'] = projectData;           

            return data;
        }
    }   
}