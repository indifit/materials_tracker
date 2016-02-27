var jw;
(function (jw) {
    (function (MaterialsTracker) {
        var Server = (function () {
            function Server() {
                this.lookupProjectFromHash = function (hash) {
                    var hashLookupSsid = jw.MaterialsTracker.Config.ConfigurationManager.getSetting(jw.MaterialsTracker.Config.ConfigurationManager.projectNumberLookupSsidKey);

                    //Open the spreadsheet using the ssid
                    var hashLookupSs = SpreadsheetApp.openById(hashLookupSsid);

                    var sheet = hashLookupSs.getSheets()[0];

                    var range = sheet.getRange(2, 1, 100, 2);

                    var projHash = jw.MaterialsTracker.Utilities.RangeUtilties.findFirstRowMatchingKey(range, hash);

                    var response = {
                        projectNumber: parseInt(projHash[1].toString()),
                        urlHash: projHash[0].toString()
                    };

                    return response;
                };
            }
            return Server;
        })();
        MaterialsTracker.Server = Server;
    })(jw.MaterialsTracker || (jw.MaterialsTracker = {}));
    var MaterialsTracker = jw.MaterialsTracker;
})(jw || (jw = {}));

var hash;

function doGet(request) {
    hash = request.parameter['projHash'];

    if (typeof hash == "undefined") {
        return HtmlService.createTemplateFromFile('InvalidProjectPage').evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }

    //Lokup the hash
    var projectLookupSsid = jw.MaterialsTracker.Config.ConfigurationManager.getSetting(jw.MaterialsTracker.Config.ConfigurationManager.projectNumberLookupSsidKey);

    var projectLookupSs = SpreadsheetApp.openById(projectLookupSsid);

    var sheet = projectLookupSs.getSheets()[0];

    var range = sheet.getRange(2, 1, 100, 2);

    var matchingRow = jw.MaterialsTracker.Utilities.RangeUtilties.findFirstRowMatchingKey(range, hash);

    if (matchingRow == null) {
        return HtmlService.createTemplateFromFile('InvalidProjectPage').evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }

    var template = HtmlService.createTemplateFromFile('Index');

    template.data = {};

    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
//# sourceMappingURL=Code.js.map
