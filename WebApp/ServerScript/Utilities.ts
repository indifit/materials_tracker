module jw.MaterialsTracker.Utilities
{
    export class RangeUtilties
    {
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

                Logger.log('keyColVal = ' + rowVals[keyColIndex]);

                if (typeof keyColVal != "undefined" && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase())
                {
                    return rowVals;
                }
            }
            return null;
        };        
    }   
}