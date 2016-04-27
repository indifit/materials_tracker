var jw;
(function (jw) {
    (function (MaterialsTracker) {
        (function (Config) {
            var ConfigurationManager = (function () {
                function ConfigurationManager() {
                }
                ConfigurationManager.configSettingsSsid = '1aHWYww4ErsnBVoxTDyjD-ZCadnW7sTaUZAt1S_n4DpQ';

                ConfigurationManager.projectNumberLookupSsidKey = 'ProjectNumberLookupSSID';

                ConfigurationManager.getSetting = function (key) {
                    //Open the config spreadsheet
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
            })();
            Config.ConfigurationManager = ConfigurationManager;
        })(MaterialsTracker.Config || (MaterialsTracker.Config = {}));
        var Config = MaterialsTracker.Config;
    })(jw.MaterialsTracker || (jw.MaterialsTracker = {}));
    var MaterialsTracker = jw.MaterialsTracker;
})(jw || (jw = {}));
//# sourceMappingURL=Config.js.map
