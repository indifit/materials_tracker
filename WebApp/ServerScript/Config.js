var jw;
(function (jw) {
    var MaterialsTracker;
    (function (MaterialsTracker) {
        var Config;
        (function (Config) {
            var ConfigurationManager = (function () {
                function ConfigurationManager() {
                }
                ConfigurationManager.configSettingsSsid = '1gBxiGSY9EjNA-ycT_JZy5C9KA0HD3Vsm9HVPs60xQws';
                ConfigurationManager.projectNumberLookupSsidKey = 'ProjectNumberLookupSSID';
                ConfigurationManager.getSetting = function (key) {
                    var configSs = null;
                    configSs = SpreadsheetApp.openById(ConfigurationManager.configSettingsSsid);
                    var configSettingsSheet = configSs.getSheetByName('ConfigSettings');
                    var configSettingsLookupRange = configSettingsSheet.getRange(1, 1, 100, 2);
                    var configSettingsMatchingRow = MaterialsTracker.Utilities.RangeUtilties
                        .findFirstRowMatchingKey(configSettingsLookupRange, key);
                    if (configSettingsMatchingRow != null) {
                        return configSettingsMatchingRow[1].toString();
                    }
                    return null;
                };
                return ConfigurationManager;
            }());
            Config.ConfigurationManager = ConfigurationManager;
        })(Config = MaterialsTracker.Config || (MaterialsTracker.Config = {}));
    })(MaterialsTracker = jw.MaterialsTracker || (jw.MaterialsTracker = {}));
})(jw || (jw = {}));
