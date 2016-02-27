var jw;
(function (jw) {
    (function (MaterialsTracker) {
        (function (Config) {
            var ConfigurationManager = (function () {
                function ConfigurationManager() {
                }
                ConfigurationManager.configSettingsSsid = '1gBxiGSY9EjNA-ycT_JZy5C9KA0HD3Vsm9HVPs60xQws';

                ConfigurationManager.projectNumberLookupSsidKey = 'ProjectNumberLookupSSID';

                ConfigurationManager.getSetting = function (key) {
                    //Open the config spreadsheet
                    var configSs = null;

                    configSs = SpreadsheetApp.openById(ConfigurationManager.configSettingsSsid);

                    var configSettingsSheet = configSs.getSheetByName('ConfigSettings');

                    var configSettingsLookupRange = configSettingsSheet.getRange(1, 1, 100, 2);

                    var configSettingsMatchingRow = MaterialsTracker.Utilities.RangeUtilties.findFirstRowMatchingKey(configSettingsLookupRange, key);

                    if (configSettingsMatchingRow != null) {
                        return configSettingsMatchingRow[1].toString();
                    }

                    return null;
                };
                return ConfigurationManager;
            })();
            Config.ConfigurationManager = ConfigurationManager;
        })(MaterialsTracker.Config || (MaterialsTracker.Config = {}));
        var Config = MaterialsTracker.Config;
    })(jw.MaterialsTracker || (jw.MaterialsTracker = {}));
    var MaterialsTracker = jw.MaterialsTracker;
})(jw || (jw = {}));
//# sourceMappingURL=Config.js.map
