module jw.MaterialsTracker.Config
{
    export class ConfigurationManager
    {
        static configSettingsSsid: string = '1aHWYww4ErsnBVoxTDyjD-ZCadnW7sTaUZAt1S_n4DpQ';

        static projectNumberLookupSsidKey: string = 'ProjectNumberLookupSSID';        

        static getSetting = (key: string): string =>
        {
            //Open the config spreadsheet
            var configSs: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ConfigurationManager.configSettingsSsid);

            var configSettingsSheet: GoogleAppsScript.Spreadsheet.Sheet = configSs.getSheetByName('MTData');

            var configSettingsLookupRange: GoogleAppsScript.Spreadsheet.Range = configSettingsSheet.getRange(2, 2, configSettingsSheet.getLastRow(), configSettingsSheet.getLastColumn());

            var configSettingsMatchingRow: Object[] = MaterialsTracker.Utilities.RangeUtilties.findFirstRowMatchingKey(configSettingsLookupRange, key);

            if (configSettingsMatchingRow != null)
            {
                return configSettingsMatchingRow[1].toString();
            }

            return null;
        };
    }
} 