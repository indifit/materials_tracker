var CDCache;
(function (CDCache) {
    function clearCache() {
        CacheService.getDocumentCache().removeAll(['CDdd', CD_CoreList, CD_WTSuppliers]);
    }
    function cacheUpdater() {
        getCentralDropDowns();
        var caches = [CD_CoreList, CD_WTSuppliers];
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();
        for (; caches.length > 0;) {
            var cacheKey = caches.pop();
            if (CacheService.getDocumentCache().get(cacheKey) == null) {
                putCache(cacheKey, cacheKey);
            }
        }
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);
    }
    function putCache(cacheKey, sheetName) {
        var _array = [];
        var _string = '';
        var _stringCaches = [];
        SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(CentralData()));
        var CD = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = CD.setActiveSheet(CD.getSheetByName(sheetName));
        _array = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
        _string = JSON.stringify(_array);
        var MSL = 100000;
        for (; _string.length > 0;) {
            if (_string.length >= MSL) {
                _stringCaches.push(_string.substr(0, MSL));
                _string = _string.substr(MSL);
            }
            if (_string.length < MSL && _string.length > 0) {
                _stringCaches.push(_string);
                _string = '';
            }
        }
        CacheService.getDocumentCache().put(cacheKey, _stringCaches.length.toString());
        for (; _stringCaches.length > 0;) {
            CacheService.getDocumentCache().put((cacheKey + _stringCaches.length), _stringCaches.pop());
        }
        return _array;
    }
    function getCache(cacheKey) {
        var cache = CacheService.getDocumentCache();
        var noOfCaches = parseInt(cache.get(cacheKey));
        var _array = [];
        var _string = '';
        if (noOfCaches != null) {
            for (var i = 1; i <= noOfCaches; i++) {
                _string += (cache.get(cacheKey + i));
            }
            _array = JSON.parse(_string);
        }
        return _array;
    }
    function getCL() {
        var _array = [];
        var _string = '';
        if (CL) {
            return CL;
        }
        var cache = CacheService.getDocumentCache();
        _string = cache.get(CD_CoreList);
        if (_string != null) {
            _array = getCache(CD_CoreList);
            CL = _array;
            return CL;
        }
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();
        _array = putCache(CD_CoreList, CD_CoreList);
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);
        return _array;
    }
    function getWTSup() {
        var _array = [];
        var _string = '';
        var cache = CacheService.getDocumentCache();
        _string = cache.get(CD_WTSuppliers);
        if (_string != null) {
            _array = getCache(CD_WTSuppliers);
            return _array;
        }
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();
        _array = putCache(CD_WTSuppliers, CD_WTSuppliers);
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);
        return _array;
    }
    function getCentralDropDowns() {
        var _object = {};
        var _string = '';
        var _CDdd = 'CDdd';
        var cache = CacheService.getDocumentCache();
        _string = cache.get(_CDdd);
        if (_string != null) {
            _object = JSON.parse(_string);
            return _object;
        }
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();
        SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(CentralData()));
        var CD = SpreadsheetApp.getActiveSpreadsheet();
        var TEMPLATE_IDS = CD.getRangeByName("TEMPLATE_IDS").getValues();
        var DRIVE_OWNER = CD.getRangeByName("DRIVE_OWNER").getValues();
        var ddType = CD.getRangeByName("dd_Type").getValues().join().split(',');
        var ddTeams = CD.getRangeByName("dd_Teams").getValues().join().split(',');
        var ddUoM = CD.getRangeByName("dd_UoM").getValues().join().split(',');
        var ddPDN = CD.getRangeByName("PD_Names").getValues().join().split(',');
        var ddPDC = CD.getRangeByName("PD_Codes").getValues().join().split(',');
        var ddVATRates = CD.getRangeByName("dd_VATRates").getValues().join().split(',');
        var ddWTSupNames = CD.getRangeByName("dd_WTSupNames").getValues().join().split(',');
        var ddStatusWT50 = CD.getRangeByName("dd_StatusWT50").getValues().join().split(',');
        var ddStatusBPR = CD.getRangeByName("dd_StatusBPR").getValues().join().split(',');
        var ddStatusHire = CD.getRangeByName("dd_StatusHire").getValues().join().split(',');
        _object = {
            DRIVE_OWNER: DRIVE_OWNER, TEMPLATE_IDS: TEMPLATE_IDS,
            ddType: ddType, ddTeams: ddTeams,
            ddUoM: ddUoM, ddPDN: ddPDN, ddPDC: ddPDC,
            ddVATRates: ddVATRates, ddWTSupNames: ddWTSupNames,
            ddStatusWT50: ddStatusWT50, ddStatusHire: ddStatusHire, ddStatusBPR: ddStatusBPR
        };
        _string = JSON.stringify(_object);
        CacheService.getDocumentCache().put(_CDdd, _string);
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);
        return _object;
    }
})(CDCache || (CDCache = {}));
