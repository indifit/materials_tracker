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

        toCamelCase = (input: string): string =>
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
            var ret: Object[] = [];

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
                    var thisPropertyName: string = this.toCamelCase(firstRow[i].toString());                    
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

            var range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(2, 1, sheet.getLastColumn(), sheet.getLastRow());

            var projHashRow: Object[] = RangeUtilties.findFirstRowMatchingKey(range, this.projectHash);

            if (projHashRow != null)
            {
                var response: jw.MaterialsTracker.Interfaces.IProjectHashLookupResponse = {
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

            data.projectHash = this.projectHash;

            return {
                templateName: templateName,
                data: data
            };
        };
        
        getIndexPageData = (projectData: MaterialsTracker.Interfaces.IProjectHashLookupResponse): Object =>
        {
            var data: any = {};

            data['projectData'] = projectData;                                                   

            return data;
        }
    }   
}