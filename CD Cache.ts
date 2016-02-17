module CDCache {
    function clearCache() {
        CacheService.getDocumentCache().removeAll(['CDdd', CD_CoreList, CD_WTSuppliers]);
    }

    function cacheUpdater() {

        getCentralDropDowns();

        var caches = [CD_CoreList, CD_WTSuppliers];

        // store current spreadsheet settings
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();

        for (; caches.length > 0;) {
            var cacheKey = caches.pop();
            if (CacheService.getDocumentCache().get(cacheKey) == null) {
                putCache(cacheKey, cacheKey);
            } // end if cache is empty
        } // end for loop

        // return to the original spreadsheet
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);

    } // end fn:cacheUpdater

    function putCache(cacheKey, sheetName) {

        var _array = [];
        var _string = '';
        var _stringCaches = []; // an array of JSON strings capped at length

        SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(CentralData()));
        var CD = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = CD.setActiveSheet(CD.getSheetByName(sheetName));
        _array = sheet.getRange(1, //starting row index
            1, // starting column index
            sheet.getLastRow(), // number of rows to import
            sheet.getLastColumn() // number of columns to import
            ).getValues(); // an array of rows, each an array of columns wholeList[r-1][c-1]

        // use the JSON.stringify method to convert the 2d array into a string for cache storage
        _string = JSON.stringify(_array);

        // check the JSON string length and break it up into smaller strings
        // !*!*!*!*!*!
        /* ! */ var MSL = 100000; // max string length
        // !*!*!*!*!*!
        for (; _string.length > 0;) {
            if (_string.length >= MSL) { // while the string is greater than or equal to the max value
                _stringCaches.push(_string.substr(0, MSL)); // split out the first MSL (e.g. 100,000) charaters
                _string = _string.substr(MSL); // write the rest of the string back
            }
            if (_string.length < MSL && _string.length > 0) {
                _stringCaches.push(_string); // write the whole remaining string
                _string = ''; // then delete the string
            }
        } // end string splitting loop

        // store the number of sub-caches in the "parent" cache
        CacheService.getDocumentCache().put(cacheKey, _stringCaches.length.toString());

        // use a loop to write the split caches
        for (; _stringCaches.length > 0;) {
            CacheService.getDocumentCache().put((cacheKey + _stringCaches.length), _stringCaches.pop());
        }

        return _array;
    } // end fn:putCache

    function getCache(cacheKey) {
        var cache = CacheService.getDocumentCache();
        var noOfCaches: number = parseInt(cache.get(cacheKey));
        var _array = [];
        var _string = '';

        if (noOfCaches != null) {

            // write the string values of the returned splpit cache into one string
            for (var i = 1; i <= noOfCaches; i++) {
                _string += (cache.get(cacheKey + i));
            }
            _array = JSON.parse(_string); // use the JSON.parse method to convert the cached string back into a 2d array
        } // if noOfCaches is not null

        return _array;

    } // end fn:getCache

    function getCL() {
        var _array = [];
        var _string = '';

        // check the "out of function" placeholder to speed up multiple calls to the cache
        if (CL) { return CL; }

        var cache = CacheService.getDocumentCache();
        _string = cache.get(CD_CoreList);
        if (_string != null) {
            _array = getCache(CD_CoreList);
            // set the "out of function" placeholder to speed up multiple calls to the cache
            CL = _array;
            return CL;
        }

        // store current spreadsheet settings
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();

        _array = putCache(CD_CoreList, CD_CoreList);

        // return to the original spreadsheet
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);

        return _array;

    } // end fn:getCL

    function getWTSup() {

        var _array = [];
        var _string = '';

        var cache = CacheService.getDocumentCache();
        _string = cache.get(CD_WTSuppliers);
        if (_string != null) {
            _array = getCache(CD_WTSuppliers);
            return _array;
        }

        // store current spreadsheet settings
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();

        _array = putCache(CD_WTSuppliers, CD_WTSuppliers);

        // return to the original spreadsheet
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);

        return _array;

    } // end fn:getWTSup

    function getCentralDropDowns() {

        var _object = {};
        var _string = '';
        var _CDdd = 'CDdd'; // the key name for the central data dropdowns


        var cache = CacheService.getDocumentCache();
        _string = cache.get(_CDdd);
        if (_string != null) {
            _object = JSON.parse(_string); // use the JSON.parse method to convert the cached string back an object of arrays
            return _object;
        }

        // store current spreadsheet settings
        var origSS = SpreadsheetApp.getActiveSpreadsheet();
        var origSH = origSS.getActiveSheet();
        var origRange = origSH.getActiveRange();

        SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(CentralData()));
        var CD: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        // the values in "" relate to the range names in the central data sheet
        // .join().split(',') is used to convert the 2d arrays into 1d arrays of string
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
            DRIVE_OWNER: DRIVE_OWNER, TEMPLATE_IDS: TEMPLATE_IDS
            , ddType: ddType, ddTeams: ddTeams
            , ddUoM: ddUoM, ddPDN: ddPDN, ddPDC: ddPDC
            , ddVATRates: ddVATRates, ddWTSupNames: ddWTSupNames
            , ddStatusWT50: ddStatusWT50, ddStatusHire: ddStatusHire, ddStatusBPR: ddStatusBPR
        };

        // use the JSON.stringify method to convert the 2d array into a string for cache storage
        _string = JSON.stringify(_object);

        // store the array in the cache
        CacheService.getDocumentCache().put(_CDdd, _string);

        // return to the original spreadsheet
        SpreadsheetApp.setActiveSpreadsheet(origSS);
        SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);

        return _object;

    } // end fn:getDropDowns  
}