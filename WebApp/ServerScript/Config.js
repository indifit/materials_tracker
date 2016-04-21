var jw;
(function (jw) {
    var MaterialsTracker;
    (function (MaterialsTracker) {
        var Config;
        (function (Config) {
            var ConfigurationManager = (function () {
                function ConfigurationManager() {
                }
                ConfigurationManager.configSettingsSsid = '1lMwQsbonNGycRoA-Zelm151YU6S_fididqPyTXDJmVQ';
                ConfigurationManager.projectNumberLookupSsidKey = 'ProjectNumberLookupSSID';
                ConfigurationManager.getSetting = function (key) {
                    var configSs = SpreadsheetApp.openById(ConfigurationManager.configSettingsSsid);
                    var configSettingsSheet = configSs.getSheetByName('MTData');
                    var configSettingsLookupRange = configSettingsSheet.getRange(2, 2, configSettingsSheet.getLastRow(), configSettingsSheet.getLastColumn());
                    var configSettingsMatchingRow = MaterialsTracker.Utilities.RangeUtilties.findFirstRowMatchingKey(configSettingsLookupRange, key);
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
