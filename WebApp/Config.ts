module jw.MaterialsTracker.Config
{
    export class ConfigurationManager
    {
        static configSettingsSsid: string = '1gBxiGSY9EjNA-ycT_JZy5C9KA0HD3Vsm9HVPs60xQws';

        static projectNumberLookupSsidKey: string = 'ProjectNumberLookupSSID';

        static getSetting = (key: string): string =>
        {
            //Open the config spreadsheet
            var configSs: GoogleAppsScript.Spreadsheet.Spreadsheet = null;

            configSs = SpreadsheetApp.openById(ConfigurationManager.configSettingsSsid);

            var configSettingsSheet: GoogleAppsScript.Spreadsheet.Sheet = configSs.getSheetByName('ConfigSettings');

            var configSettingsLookupRange: GoogleAppsScript.Spreadsheet.Range = configSettingsSheet.getRange(1, 1, 100, 2);

            var configSettingsMatchingRow: Object[] =
                MaterialsTracker.Utilities.RangeUtilties
                    .findFirstRowMatchingKey(configSettingsLookupRange,
                    key);

            if (configSettingsMatchingRow != null)
            {
                return configSettingsMatchingRow[1].toString();
            }

            return null;
        };
    }
} 