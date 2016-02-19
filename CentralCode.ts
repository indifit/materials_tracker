 ﻿/**
 * Functions in this code file are:
 * onOpen
 * purchasingSidebar
 * onInstall
 * edited
 * clearCache
 * cacheUpdater
 * putCache
 * getCache
 * getCL
 * getWTSup
 * getCentralDropDowns
 * setupCLPicker
 * clEdit
 * setCLFilters
 * displayFilteredCLList
 * checkForDuplicates
 * filterCLList
 * setCLDropDown
 * basketList
 * sendBasket
 * searchCLbyIC
 * clearCells
 * setupGoodsReceiver
 * getML
 * grEdit
 * setGRFilters
 * clearDisplayGRList
 * compileFilteredGRList
 * filterGRList
 * setGRDropDown
 * markGoodsReceived
 * matEdit
 * matCleanup
 * matDropDowns
 * matDateFormat
 * matNumberFormat
 * matSetFormulae
 * col
 * updateStatustoSent
 * ncEdit
 * processNonCore
 * POs3_createPO
 * nextPOPRNumber
 * getSupplier
 * POs4_updateMaterialsList
 * generatePO
 * POs1_getItems
 * confirmGeneratePO
 * POs2_handleIncompleteItems
 * PRs3_createPR
 * PRs4_updateMaterialsList
 * PR_Email
 * getAsExcel
 * svsEdit
 * svsTester
 * getPRDataReadyForSVS
 * confirmMatchedItems
 * getPDFSVS
 */

/*
 * Global scope variables
 */
var UI = null;
var CL = null; //out of script holder for the CL return to speed up cache calls

var clHeaderRow = 6;
var ddPDN = 'C2';
var ddTrade = 'C3';
var ddSubcategory = 'D3';
var ddType = 'E3';
var ddLocation = 'H3';
var ddPackage = 'C4'; /* WIP */
var ddTeamRequested = 'G3';
var ddSendBasket = 'J3';
var ddViewAll = 'A2';

var dupWarningCell = 'A4';
var dupResultsCell = 'B4';

// sort options
var clSortRow = 5;

// headers as in Central Data Core List for filtering
var clSubcategory = 'Product sub category';
var clType = 'Type';
var clLocation = 'Location';
var clPackage = 'Package key';
var clPDC = 'PDC';

// other headers from Central Data Core List
var clTrade = 'Trade';
var clIC = 'Item Code';
var clID = 'Item Description';
var clLastPrice = 'Expected Purchase Price';
var clPUoM = 'Purchase UOM';
var clFactor = 'Factor';
var clBUoM = 'Base UOM';
var clBrand = 'Brand';
var clMfg = 'Manufacturer';
var clPartNo = 'Part Number';
var clItemStatus = 'Status';
var clInactive = 'Inactive from';
var clReplacement = 'Replacement';
var clDataLink = 'Data Link';

// headers from basket
var bIC = 'Item Code';
var bID = 'Item Description';
var bMfg = 'Manufacturer';
var bBrand = 'Brand';
var bPartNo = 'Part Number';
var bLastPrice = 'Last Price';
var bQty = 'Quantity Requested';
var bPUoM = 'Purchase UoM';
var bFactor = 'Factor';
var bBUoM = 'Base UoM';
var bPDN = 'Budget PDC';
var bUsage = 'Usage';

var grHeaderRow = 6;
var grSortRow = 5;
var ddPurchasingRoute = 'F3';
var supplierColNum = 6;
var mfgColNum = 4;
var partColNum = 5;
var grLastUpdated = 'G3';
var ddMarkReceived = 'L3';

var grCOREopt = 'Core';
var grWTopt = WT_PREFIX;
var grSHOWALLopt = 'GO!';

var MATLIST = null;

var hRow = null;

var ncHeaderRow = 2;
var ddProcessNonCore = 'G2';

// single values for PO entries from materials list of thisRow
var poSupplierName = '';
var poDelivery = '';
var poNumber = '';
var requestingDept = '';
var sameDept = '';
var delONBY = '';

// multi dimensional array holder for the whole materials sheet
var wholeRange = '';
var wholeList = [];
// multi dimensional array holder for the wt suppliers import
var aCD_WTSuppliers = [];

// arrays for PO entries from materials list
var iQty = [];
var iUoM = [];
var iFactor = [];
var iBUoM = [];
var iCode = []; // item code used separately for PRs
var iDesc = []; // item description is combined with [item code] for Materials and Hire. Also combined with on delivery / off delivery for hire items
var iUnit = [];
var iPDN = []; // project dimension name, not actually on PO, but place in hidden column
var iPDC = []; // project dimension code, used for BPR
var iIndices = []; // indices numbers for line items to have a PO added
var iNote = [];

var aRegenIndices = []; // array of indices to remove regen links

var isEmergency = false;
var branchPONumber = '';

var svsSH = SS.getSheetByName(SVS_MATCHER_SHEET);

var svsSH_HR = 6; // the header row for the svs matcher sheet
var svsSortRow = 5;

var ddSVSProcess = 'B2';
var ddSVSMatchConfirm = 'H2';
var svsLastUpdated = 'C3';


/* SVS MATCHER COLUMN HEADERS */

var svsMAT_PR = 'PR Number';
var svsMAT_LINE = 'Line ID';
var svsMAT_DESC = 'Item Description';
var svsMAT_QTY = 'Item Quantity';
var svsMAT_DEL = 'Requested Delivery Date';

var svsSVS_SUPPLIER = 'SVS Supplier';
var svsSVS_WTPO_NUM = 'SVS WT PO Number';
var svsSVS_WTPO_LINE = 'SVS WT PO Line';

/*
 * End global scope variables
 */

/*
* Functions originally in Code.gs
*/

function onOpen(e) {
    UI = SpreadsheetApp.getUi();

    var statusUpdateSubMenu =
        UI.createMenu('Update Order Status')
            .addItem('...to Sent', 'updateStatustoSent')

  UI.createMenu('Purchasing Tools')
        .addItem('Show Purchasing Sidebar', 'purchasingSidebar')
        .addSeparator()
        .addSubMenu(statusUpdateSubMenu)
        .addSeparator()
        .addItem('Clean Up Formats, Formulae and Drop-downs', 'matCleanup')
        .addToUi();
    setupCLPicker();
    setupGoodsReceiver();
    matDropDowns();
}

/*
* Settings for the purchasing side bar
*/
function purchasingSidebar() {
    UI = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutputFromFile('SideBarPlain')
        .setSandboxMode(HtmlService.SandboxMode)
        .setTitle('Purchasing Side Bar')
        .setWidth(250);
    UI.showSidebar(html);
}

/**
 * Runs when the add-on is installed; calls opened() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
    onOpen(e);
}

function edited(e) {    
    UI = SpreadsheetApp.getUi();
    SH = SS.getActiveSheet();
    var SHname = SH.getName();
    thisCell = SH.getActiveCell();

    if (SHname == CORELIST_SHEET) {
        clEdit(e);
    }

    if (SHname == MATERIALS_SHEET) {
        matEdit(e);
    }

    if (SHname == GOODS_RECEIVING_SHEET) {
        grEdit(e);
    }

    if (SHname == SVS_MATCHER_SHEET) {
        svsEdit(e);
    }

    if (SHname == NONCORE_SHEET) {
        ncEdit(e);
    }

} // end fn:onEdit

/*
 * End function originally in Code.gs
 */


/*
 * Functions originally in CD Cache.gs
 */
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

    var _object: MaterialsTrackerInterfaces.ICentralDropDowns;
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
        DRIVE_OWNER: DRIVE_OWNER,
        TEMPLATE_IDS: TEMPLATE_IDS,
        ddType: ddType,
        ddTeams: ddTeams,
        ddUoM: ddUoM,
        ddPDN: ddPDN,
        ddPDC: ddPDC,
        ddVATRates: ddVATRates,
        ddWTSupNames: ddWTSupNames,
        ddStatusWT50: ddStatusWT50,
        ddStatusHire: ddStatusHire,
        ddStatusBPR: ddStatusBPR
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

/*
 * End functions originally in CD Cache.gs
 */

/*
 * Functions originally in CoreList.gs
 */

function setupCLPicker() {

    getCL();
    var CDD = getCentralDropDowns();
    /*  setCLDropDown(CL,ddPackage,clPackage);
    */


    CDD.ddPDN.pop(); // remove the last (void) entry

    var cells = [ // an array of columnHeader and dropDown datasource pairs
        ddPDN, CDD.ddPDN
        , ddTeamRequested, CDD.ddTeams
    ];

    for (; cells.length > 0;) {
        var cell = cells.shift(); // the cell reference (first of the array pair)
        var dv = SpreadsheetApp.newDataValidation();
        dv.setAllowInvalid(false);
        var _dd = cells.shift(); // the dropdown source datalist, (second of the array pair)
        dv.requireValueInList(_dd, true);
        SS.getSheetByName(CORELIST_SHEET).getRange(cell).setDataValidation(dv.build());
    } // end for loop

}

function clEdit(e) {
    var eRangeA1 = e.range.getA1Notation();
    var eValues = e.range.getValues();

    if ((eRangeA1 == ddPDN) || (eRangeA1 == ddTrade) || (eRangeA1 == ddSubcategory) || (eRangeA1 == ddType) || (eRangeA1 == ddViewAll)) {
        SS.toast("Please wait...", "Working", 600);
        setCLFilters(e);
    }

    if (eRangeA1 == ddSendBasket && e.value == 'GO!') {
        SS.toast("Please wait...", "Working", 600);
        sendBasket(e);
    }

    if (e.range.getRow() == clSortRow) {

        var sortDir = null;

        if (eValues[0][0] == 'A > Z') { sortDir = true; }
        if (eValues[0][0] == 'Z > A') { sortDir = false; }

        if (sortDir != null) {
            SH.getRange((clHeaderRow + 1), //starting row index
                1, // starting column index
                SH.getLastRow(), // number of rows to import
                SH.getLastColumn() // number of columns to import
                ).sort({ column: e.range.getColumn(), ascending: sortDir });
        }// end if sort is valid
        e.range.setValue('Sort?');
    } // end if sort row was edited

    if (e.range.getRow() > clHeaderRow) {

        if (e.range.getColumn() == 7) { // check if the quantity column was edited
            var _PDN = SH.getRange(ddPDN).getValue();
            // the budget range is the same size as the edited range, but 2 columns to the right
            var budgetPDNRange = e.range.offset(0, 4);
            var budgetPDNValues = budgetPDNRange.getValues();

            for (var i = 0; i < eValues.length; i++) {
                if (eValues[i][0] > 0 && budgetPDNValues[i][0] == '') {
                    budgetPDNValues[i][0] = _PDN;
                } // end if a quantity was picked and the item didn't already have a pdn
                if (!eValues[i][0]) {
                    budgetPDNValues[i][0] = '';
                } // end if the quanity was set to 0 and so the pdn was cleared
            } // end loop through edited quantity values vs budget pdn

            budgetPDNRange.setValues(budgetPDNValues);

        } // end if quantity column was edited

        // grab everything, including the basket header
        var currentFLRange = SH.getRange(clHeaderRow, //starting row index
            1, // starting column index
            SH.getLastRow(), // number of rows to import
            SH.getLastColumn() // number of columns to import
            );

        var basketReturn = basketList(currentFLRange.getValues());

        var basket = basketReturn.basket; // an array for the whole basket (every item with qty>0)
        var basketICs = basketReturn.ics; // an array of just the item codes (as strings) of the items in the basket
        var basketDLs = basketReturn.dls; // an array of just the item codes data links (as strings) of the items in the basket

        var duplicates = checkForDuplicates(basketICs);

        if (duplicates) {
            SH.getRange(dupWarningCell).setValue('Duplicates\nFound');
            SH.getRange(dupResultsCell).setValue(duplicates.join('\n'));
        } // end if duplicates are found
    } // end if basket area was edited

} // end fn:clEdit

// e is the event object from the edit event in Code.gs
function setCLFilters(e) {

    var cellRef = e.range.getA1Notation();
    var cellValue = e.value;
    if (!cellValue) { cellValue = ''; }

    var fl = []; // the filtered core list

    /*  var packageVal = SH.getRange(ddPackage).getValue();
      if (!packageVal){packageVal='';}
      var flPackage = filterCLList(getCL(),packageVal,clPackage);
    */

    var viewAllVal = SH.getRange(ddViewAll).getValue();

    var PDCVal = PD(SH.getRange(ddPDN).getValue());
    if (!PDCVal) { PDCVal = ''; }
    var flPDC = filterCLList(getCL(), PDCVal, clPDC);
    /*  // if a package has been chosen, limit the PDC selection to those items from that package,
      // ...which will cascade to sub and type
      if (packageVal.length>0){flPDC = filterCLList(flPackage,PDNVal,clPDC);}
    */

    var tradeVal = SH.getRange(ddTrade).getValue();
    var flTrade = filterCLList(flPDC, tradeVal, clTrade);

    var subVal = SH.getRange(ddSubcategory).getValue();
    var flSub = filterCLList(flTrade, subVal, clSubcategory);

    var typeVal = SH.getRange(ddType).getValue();
    var flType = filterCLList(flSub, typeVal, clType);

    /*  // >>>> Package dropdown was edited ***************** WIP
      // check which other filters are in place
      if (cellRef == ddPackage){
        if (cellValue.length>0){fl = flPackage;Logger.log("package");} // end if cellValue was present
        if (PDNVal.length>0){fl = flPDC;Logger.log("pdc");} // unnecessary line of code, but there for logical consitency
        if (subVal.length>0){fl = flSub;Logger.log("sub");}
        if (typeVal.length>0){fl = flType;Logger.log("type");}
        setCLDropDown(flPDC,ddPDN,clPDC);
        setCLDropDown(flSub,ddSubcategory,clSubcategory);
        setCLDropDown(flType,ddType,clType);
      } // end if package
    */
    // >>>> PDC dropdown was edited
    if (cellRef == ddPDN) {
        clearCells([ddTrade, ddSubcategory, ddType]);
        if (cellValue.length > 0) {
            fl = flPDC;
            setCLDropDown(fl, ddTrade, clTrade);
        } // end if cellValue was present
    } // end if PDC

    // >>>> Trade dropdown was edited
    if (cellRef == ddTrade) {
        clearCells([ddSubcategory, ddType]);
        if (cellValue.length > 0) {
            fl = flTrade;
            setCLDropDown(fl, ddSubcategory, clSubcategory);
        } // end if cellValue was present
    } // end if Trade

    // >>>> Subcategory dropdown was edited
    if (cellRef == ddSubcategory) {
        clearCells([ddType]);
        if (cellValue.length > 0) {
            fl = flSub;
            setCLDropDown(fl, ddType, clType);
        } // end if cellValue was present
        if (cellValue.length == 0) {
            fl = flTrade;
            //      setCLDropDown(fl,ddType,clType);
        } // reset to flPDC level of filter if subcategory was deleted
    } // end if subcategory

    // >>>> Type dropdown was edited
    if (cellRef == ddType) {
        if (cellValue.length > 0) {
            fl = flType;
        } // end if cellValue was present
        if (cellValue.length == 0) {
            fl = flSub;
        } // reset to flSub level of filter if type was deleted
    } // end if type

    // >>>> ViewAll dropdown was edited
    if (cellRef == ddViewAll) {
        if (cellValue == 'View All') {
            clearCells([ddPDN, ddTrade, ddSubcategory, ddType]);
            fl = getCL();
        } // end if cellValue was to view all
        if (cellValue != 'View All') {
            setupCLPicker();
        }
    } // end if ViewAll

    // filter the CL
    displayFilteredCLList(fl);
    SS.toast("Thanks for waiting", "Done", 1);
} // fn:setCLFilters

function displayFilteredCLList(fl) {

    // grab everything, including the basket header
    var currentFLRange = SH.getRange(clHeaderRow, //starting row index
        1, // starting column index
        SH.getLastRow(), // number of rows to import
        SH.getLastColumn() // number of columns to import
        );

    var basketReturn = basketList(currentFLRange.getValues());

    var basket = basketReturn.basket; // an array for the whole basket (every item with qty>0)
    var basketICs = basketReturn.ics; // an array of just the item codes (as strings) of the items in the basket
    var basketDLs = basketReturn.dls; // an array of just the item codes data links (as strings) of the items in the basket
    var basketPDNs = basketReturn.pdns; // an array of just the project dimension names (as strings) of the items in the basket

    var duplicates = checkForDuplicates(basketICs);

    if (duplicates) {
        SH.getRange(dupWarningCell).setValue('Duplicates\nFound');
        SH.getRange(dupResultsCell).setValue(duplicates.join('\n'));
    }


    // clear the basket area, except for the header
    SH.getRange((clHeaderRow + 1), //starting row index
        1, // starting column index
        SH.getLastRow(), // number of rows to import
        SH.getLastColumn() // number of columns to import
        ).clearContent();

    // dl is the display list. This is fl with filter columns (trade, sub, type, etc.) removed
    var dl = [];
    var dlDataLinks = [];

    var _PDN = SH.getRange(ddPDN).getValue();

    // start var row at 1 to ignore header row of fl
    for (var row = 1; row < fl.length; row++) {

        var _IC = fl[row][fl[0].indexOf(clIC)];
        if (basketICs.indexOf(_IC) < 0 // check that the item code for this display item isn't in the basket
        // or if it is, that it is from a PDN not already used
            || (basketICs.indexOf(_IC) >= 0 && basketPDNs.indexOf(_PDN) < 0)
            ) {
            var dlItem = [];
            var dlDataLink = [];
            // push in the order needed for the basket columns
            dlItem.push(_IC);
            dlItem.push(fl[row][fl[0].indexOf(clID)]);
            dlItem.push(fl[row][fl[0].indexOf(clMfg)]);
            dlItem.push(fl[row][fl[0].indexOf(clBrand)]);
            dlItem.push(fl[row][fl[0].indexOf(clPartNo)]);
            dlItem.push(fl[row][fl[0].indexOf(clLastPrice)]);
            dlItem.push(''); // a blank entry for the quantity requested column
            dlItem.push(fl[row][fl[0].indexOf(clPUoM)]);
            dlItem.push(fl[row][fl[0].indexOf(clFactor)]);
            dlItem.push(fl[row][fl[0].indexOf(clBUoM)]);

            dl.push(dlItem);

            // get the data link, if available
            var _dataLink = fl[row][fl[0].indexOf(clDataLink)];
            dlDataLink.push('=HYPERLINK(\"' + _dataLink + '\",\"' + _IC + '\")');
            dlDataLinks.push(dlDataLink);
        } // end if dl item is not already in the basket
    } // end for loop

    basket.shift(); // remove the header row

    if (basket.length > 0) {
        // get a range equal in row length to the basket. Write even if there is no basket so that the header is written back in.
        SH.getRange((clHeaderRow + 1), //starting row index
            1, // starting column index
            basket.length, // number of rows to import
            basket[0].length // number of columns to import
            ).setValues(basket); // set this range with the contents of the filtered list.

        SH.getRange((clHeaderRow + 1), //starting row index, just below header row
            1, // starting column index
            basketDLs.length, // number of rows to import
            1 // number of columns to import
            ).setFormulas(basketDLs); // set this range with the formula contents of the filtered list.
    }

    // only export the filtered display list if one exists
    if (dl.length > 0) {
        // get a range equal in row length to the filtered list and start below the existing basket
        SH.getRange((clHeaderRow + 1 + basket.length), //starting row index
            1, // starting column index
            dl.length, // number of rows to import
            dl[0].length // number of columns to import
            ).setValues(dl); // set this range with the contents of the filtered list.

        // write the data links into the first column
        SH.getRange((clHeaderRow + 1 + basket.length), //starting row index
            1, // starting column index
            dlDataLinks.length, // number of rows to import
            1 // number of columns to import
            ).setFormulas(dlDataLinks); // set this range with the formula contents of the filtered list.

    } // end if display list has content / exists

    // auto resize all the basket columns
    //for (var c=1;c<9;c++){SH.autoResizeColumn(c);}
    SH.getRange(1, 1, 1000, 1).setNumberFormat('@STRING@'); // set the first column to plain text
    //  SH.autoResizeColumn(2);

} // end fn:displayFilteredList

function checkForDuplicates(basketICs) {

    SH.getRange(dupWarningCell + ':' + dupResultsCell).clearContent();

    // check for duplicated items already on the materials list
    var tempMatSh = SS.getSheetByName(MATERIALS_SHEET);
    var tempMatList = tempMatSh.getRange(1, 1, tempMatSh.getLastRow(), 9).getValues();

    var duplicatesArray = [];

    for (var t = (tempMatList.length - 1); t > 0; t--) {
        // if this iteration of the materials list has an item code that matches anything in the basket...
        var thisIC = tempMatList[t][tempMatList[MHRI].indexOf(H_ICODE)];
        var thisICasString = thisIC.toString();
        if (basketICs.indexOf(thisIC) > -1 || basketICs.indexOf(thisICasString) > -1) {
            //... insert tab separated row number, line id, item code and existing quantity into the duplicates array
            var dA = [
                (t + 1),
                (tempMatList[t][tempMatList[MHRI].indexOf(H_LINEID)]),
                thisIC,
                (tempMatList[t][tempMatList[MHRI].indexOf(H_QTY)])
            ];
            duplicatesArray.unshift(dA.join('\t'));
        } // end if duplicate has been found
    } // end for loop

    // if duplicates were found, insert a header "row"
    if (duplicatesArray.length > 0) {
        var _dA = ['Row', 'Line', 'Item', 'Qty'];
        duplicatesArray.unshift(_dA.join('\t'));
        return duplicatesArray;
    }

    return null;

} /// end fn:checkForDuplicates

function filterCLList(list, ddValue, ddOption) {
    var subList = [];
    // for the provided List, find all items for this dropDown option
    for (var i = 1; i < list.length; i++) {
        var li = list[i][list[0].indexOf(ddOption)].toString();//store the list item as a string
        if (li.indexOf(ddValue) > -1) { //check if this item's option field contains a string match to the search value
            subList.push(list[i]);
        } // end if
    } // end for loop
    subList.unshift(getCL()[0]);// add in the title row at the beginning
    return subList;
} // end fn:filterList

function setCLDropDown(list, ddCell_A1, ddType) {
    var aOpt = []; // an array of the dropdown options
    // for the provided List, find all available options for this Type. Start at i=1 to remove header
    // the header is unshifted back to the top of the array at the end of fn:filterList
    for (var i = 1; i < list.length; i++) {
        var opt = list[i][list[0].indexOf(ddType)];
        if (aOpt.indexOf(opt) < 0) {
            aOpt.push(opt); // if this type isn't already in the dropdown option list, add it
        } // end if
    } // end for loop
    aOpt.sort(); // sort alphabetically

    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.requireValueInList(aOpt, true);
    SS.getSheetByName(CORELIST_SHEET).getRange(ddCell_A1).setDataValidation(dv.build());

} // end fn:setCLDropDown

function basketList(list) {
    var subList = [];
    var icList = [];
    var dlList = [];
    var pdnList = [];
    // for the provided List, find all items with a quantity greater than one
    for (var i = 1; i < list.length; i++) {
        if (list[i][list[0].indexOf(bQty)] > 0) {
            subList.push(list[i]); // the whole item
            pdnList.push(list[i][list[0].indexOf(bPDN)]);
            var _ic = list[i][list[0].indexOf(bIC)].toString();
            icList.push(_ic); // the item code as a string
            var _dataLink = searchCLbyIC(_ic, clDataLink);
            var _dlFormula = [];
            _dlFormula.push('=HYPERLINK(\"' + _dataLink + '\",\"' + _ic + '\")');
            dlList.push(_dlFormula);
        } // end if
    } // end for loop
    subList.unshift(list[0]);// add in the title row at the beginning
    var retVal = { basket: subList, ics: icList, dls: dlList, pdns: pdnList };
    return retVal;
} // end function

function sendBasket(e) {

    // store the last column index value so that it's not lost after range.clear()
    var clLastCol = SH.getLastColumn();

    // grab everything, including the basket header
    var currentFLRange = SH.getRange(clHeaderRow, //starting row index
        1, // starting column index
        SH.getLastRow(), // number of rows to import
        clLastCol // number of columns to import
        );

    var basketReturn = basketList(currentFLRange.getValues());

    var basket = basketReturn.basket; // an array for the whole basket (every item with qty>0)
    var basketICs = basketReturn.ics; // an array of just the item codes (as strings) of the items in the basket
    var basketDLs = basketReturn.dls; // an array of just the item codes data links (as strings) of the items in the basket

    var duplicates = checkForDuplicates(basketICs);

    Logger.log('tested for duplicates');

    /*  if (duplicates){

        Logger.log('duplicates found');

        var duplicatesAlert = UI.alert("Item Duplication",
                                       "Item(s) in your Core List Basket are already present on the Materials List."
                                       +"\nA summary of any duplication is shown at the top left of the Core List Picker."
                                       +"\nChoose \"OK\" to continue and add your basket to the Materials List, or \"Cancel\" to stop and resolve the duplication.",
                                       UI.ButtonSet.OK_CANCEL);

        if (duplicatesAlert == UI.Button.OK){
          Logger.log('said yes to process duplicates');
        }

        if (duplicatesAlert == UI.Button.CANCEL){

          SH.getRange(dupWarningCell).setValue('Duplicates\nFound');
          SH.getRange(dupResultsCell).setValue(duplicates.join('\n'));

          e.range.setValue("Add to Materials Tracker");
          SS.toast(".", ".", 1);
          return;
        }

      } // end if duplicates were found
    */
    Logger.log('basket size is: ' + (basket.length - 1));

    if (basket.length == 1 && e.value == 'GO!') {
        e.range.setValue("Add to Materials Tracker");
        SS.toast(".", ".", 1);
        UI.alert("No Items to Add", "There were no basket items found to add to the materials tracker", UI.ButtonSet.OK);
        return;
    }

    currentFLRange.clearContent(); // clear the filter view

    var basketHeader = SH.getRange(clHeaderRow, 1, 1, clLastCol);
    basketHeader.setValues([basket[0]]);

    _sh = SS.getSheetByName(MATERIALS_SHEET)
  var curMatRange = _sh.getRange(1, //starting row index
        2, // starting column index
        _sh.getLastRow(), // number of rows to import
        _sh.getLastColumn() // number of columns to import
        );

    var curMat = curMatRange.getValues();
    var blankRow = 0;

    for (var i = curMat.length - 1; i > MHRI; i--) {
        // check for the first blank row
        if (
            (curMat[i][curMat[MHRI].indexOf(H_IDESC)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_ICODE)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_TYPE)] != '')
            ) { blankRow = i + 2; break; }
    } // end for loop

    // basketExport (be) is the range [r][c] of ordered columns for the materials list
    var be = [];
    var bePDN = [];
    var beSupplier = [];
    var beStatus = [];
    // start var r at 1 to ignore header row of fl
    for (var b = 1; b < basket.length; b++) {
        var beItem = [];
        var clType = CORELIST;
        // push in the order needed for the basket columns to match the materials tracking sheet
        beItem.push(basket[b][basket[0].indexOf(bID)]);
        beItem.push(basket[b][basket[0].indexOf(bIC)]);
        if (beItem[1].indexOf('(') > -1) { clType = COREEXTRA; } // parenthesis in the item code indicates Free Text item which is a Core Extra
        beItem.push(clType);
        beItem.push(''); // a blank entry for the notes column
        beItem.push(basket[b][basket[0].indexOf(bUsage)]);
        beItem.push(SH.getRange(ddTeamRequested).getValue());
        beItem.push(basket[b][basket[0].indexOf(bQty)]);
        beItem.push(basket[b][basket[0].indexOf(bPUoM)]);
        beItem.push(basket[b][basket[0].indexOf(bFactor)]);
        beItem.push(basket[b][basket[0].indexOf(bBUoM)]);
        beItem.push(basket[b][basket[0].indexOf(bLastPrice)]);
        be.push(beItem);

        var _bePDN = [];
        _bePDN.push(basket[b][basket[0].indexOf(bPDN)]);
        bePDN.push(_bePDN)

    var _beSupplier = [];
        _beSupplier.push(PR_SUPPLIER)
    beSupplier.push(_beSupplier);

        var _beStatus = [];
        _beStatus.push('1 To Be Reviewed');
        beStatus.push(_beStatus);

    } // end for loop

    _sh.getRange(blankRow, //starting row index
        col(H_IDESC).n, // starting column index
        be.length, // number of rows to import
        be[0].length // number of columns to import
        ).setValues(be);

    _sh.getRange(blankRow, //starting row index
        col(H_PDN).n, // starting column index
        bePDN.length, // number of rows to import
        bePDN[0].length // number of columns to import
        ).setValues(bePDN);

    _sh.getRange(blankRow, //starting row index
        col(H_SUPPLIER).n, // starting column index
        beSupplier.length, // number of rows to import
        beSupplier[0].length // number of columns to import
        ).setValues(beSupplier);

    _sh.getRange(blankRow, //starting row index
        col(H_STATUS).n, // starting column index
        beStatus.length, // number of rows to import
        beStatus[0].length // number of columns to import
        ).setValues(beStatus);

    // set the "completed" notification, clear the filter dropdowns and reset them.
    e.range.setValue("Add to Materials Tracker");
    clearCells([ddPDN, ddTrade, ddSubcategory, ddType]);
    setupCLPicker();
    SS.toast("Thanks for waiting", "Done", 1);
}

function searchCLbyIC(IC, searchHeader) {
    CL = getCL();
    var retVal = null;
    for (var i = 0; i < CL.length; i++) {
        // search through for the IC
        if (CL[i][CL[0].indexOf(clIC)] == IC) {
            retVal = CL[i][CL[0].indexOf(searchHeader)];
            break;
        } // end if ic found
    } // end for loop
    return retVal;
} // end fn:searchCLbyIC

function clearCells(cellsA1) {
    for (; cellsA1.length > 0;) {
        SH.getRange(cellsA1.pop()).clearContent().clearDataValidations();
    }
} // end fn:clearCells

/*
 * End functions originally in CoreList.gs
 */

/*
 * Functions originally in GoodsReceiving.gs

 */
/*
* some functionality is shared with the core list
* particularly dropdown filtering and display.
* however, to maintain independance, functions are duplicated
*/

function setupGoodsReceiver() {
    /*
    var CDD = getCentralDropDowns();

    var cells = [ // an array of columnHeader and dropDown datasource pairs
      ddPurchasingRoute , [grSHOWALLopt,grCOREopt,grWTopt]
    ];

    for (;cells.length>0;){
      var cell = cells.shift(); // the cell reference (first of the array pair)
      var dv = SpreadsheetApp.newDataValidation();
      dv.setAllowInvalid(false);
      var _dd = cells.shift(); // the dropdown source datalist, (second of the array pair)
      dv.requireValueInList(_dd, true);
      SS.getSheetByName(GOODS_RECEIVING_SHEET).getRange(cell).setDataValidation(dv.build());
    } // end for loop
  */
}

function getML() {
    if (!MATLIST) {
        MATLIST = matSH.getRange(
            (MHRI + 1), 1, matSH.getLastRow(), matSH.getLastColumn()
            ).getValues();
    }
    return MATLIST;
} // end fn:getML

function grEdit(e) {
    var eRangeA1 = e.range.getA1Notation();
    var eValues = e.range.getValues();

    if (eRangeA1 == ddPurchasingRoute && e.value == 'GO!') {
        SS.toast("Please wait...", "Working", 600);
        setGRFilters();
    }

    if (eRangeA1 == ddMarkReceived && e.value == 'GO!') {
        SS.toast("Please wait...", "Working", 600);
        markGoodsReceived();
    }

    if (e.range.getRow() == grSortRow) {

        var sortDir = null;

        if (eValues[0][0] == 'A > Z') { sortDir = true; }
        if (eValues[0][0] == 'Z > A') { sortDir = false; }

        if (sortDir != null) {
            SH.getRange((clHeaderRow + 1), //starting row index
                1, // starting column index
                SH.getLastRow(), // number of rows to import
                SH.getLastColumn() // number of columns to import
                ).sort({ column: e.range.getColumn(), ascending: sortDir });
        }// end if sort is valid
        e.range.setValue('Sort?');
    } // end if sort row was edited

    if (e.range.getRow() > grHeaderRow) {

        if (e.range.getColumn() == 12) { // check if the quantity column was edited
            var thisDate = new Date();
            // the budget range is the same size as the edited range, but 2 columns to the right
            var dateRange = e.range.offset(0, 1);
            var dateValues = dateRange.getValues();

            for (var i = 0; i < eValues.length; i++) {
                if (eValues[i][0] > 0 && dateValues[i][0] == '') {
                    dateValues[i][0] = thisDate;
                } // end if a quantity was picked and the item didn't already have a pdn
                if (!eValues[i][0]) {
                    dateValues[i][0] = '';
                } // end if the quanity was set to 0 and so the pdn was cleared
            } // end loop through edited quantity values vs budget pdn

            dateRange.setValues(dateValues);

        } // end if quantity column was edited

    } // end if basket area was edited

} // end fn:grEdit

// e is the event object from the edit event in Code.gs
function setGRFilters() {

    var fl = []; // the filtered list

    var purchasingRouteVal = SH.getRange(ddPurchasingRoute).getValue();
    if (!purchasingRouteVal) { purchasingRouteVal = ''; }
    var flPurchasingRoute = filterGRList(getML(), purchasingRouteVal, H_TYPE);
    if (purchasingRouteVal == grSHOWALLopt) {
        var flCore = filterGRList(getML(), grCOREopt, H_TYPE);
        var flWT = filterGRList(getML(), grWTopt, H_TYPE);
        flWT.shift(); // remove the first entry which is the header
        fl = flCore.concat(flWT);
    }

    Logger.log(flPurchasingRoute);

    /*
      var supplierVal = SH.getRange(ddSupplier).getValue();
      var supplierColHeader = H_BRANCHSUPPLIER;
      if (purchasingRouteVal == grWTopt){supplierColHeader = H_SUPPLIER;}
      if (!supplierVal){supplierVal = '';}
      var flSupplier = filterGRList(flPurchasingRoute,supplierVal,supplierColHeader);

      // date . . . needs work
    //  var typeVal = SH.getRange(ddType).getValue();
    //  var flType = filterList(flSub,typeVal,clType);

      // >>>> Purchasing Route dropdown was edited
      if (cellRef == ddPurchasingRoute){
        clearCells([ddSupplier]);
        SH.showColumns(supplierColNum);
        if (cellValue.length>0){
          fl = flPurchasingRoute;
          setGRDropDown(fl,ddSupplier,supplierColHeader);
        } // end if cellValue was present
        if (cellValue == grWTopt){
          SH.hideColumns(mfgColNum);
          SH.hideColumns(partColNum);
        }
        if (cellValue == grCOREopt || cellValue == grSHOWALLopt){
          SH.showColumns(mfgColNum);
          SH.showColumns(partColNum);
        }
      } // end if Purchasing Route

      // >>>> Supplier dropdown was edited
      if (cellRef == ddSupplier){
        if (cellValue.length>0){
          fl = flSupplier;
          SH.hideColumns(supplierColNum);
        } // end if cellValue was present
        if (cellValue.length==0){
          fl=flPurchasingRoute;
          SH.showColumns(supplierColNum);
        } // reset to flPurchasingRoute level of filter if subcategory was deleted
      } // end if Supplier
    */
    // filter the GR List
    clearDisplayGRList(); // clear the currently displayed list

    var isCore = true;
    //  if (purchasingRouteVal == grWTopt){isCore=false;}
    //  if (purchasingRouteVal != grSHOWALLopt){compileFilteredGRList(fl,isCore);}

    compileFilteredGRList(fl, true);

    SH.getRange(grLastUpdated).setValue(new Date());
    SH.getRange(ddPurchasingRoute).setValue('Refresh Items');

    SH.getRange((clHeaderRow + 1), //starting row index
        1, // starting column index
        SH.getLastRow(), // number of rows to import
        SH.getLastColumn() // number of columns to import
        ).sort({ column: 8, ascending: true });

    SS.toast("Thanks for waiting", "Done", 1);

} // fn:setGRFilters

function clearDisplayGRList() {
    // grab everything below the header
    var currentFLRange = SH.getRange((grHeaderRow + 1), //starting row index
        1, // starting column index
        SH.getLastRow(), // number of rows to import
        SH.getLastColumn() // number of columns to import
        );

    var currentFLValues = currentFLRange.getValues();

    // check if anything has been marked goods received

    currentFLRange.clearContent(); // clear the filter view
} // end fn:clearDisplayGRList

function compileFilteredGRList(fl, isCore) {

    // dl is the display list. This is fl with filter columns removed
    var dl = [];
    var dlDataLinks = [];

    // start var row at 1 to ignore header row of fl
    for (var row = 1; row < fl.length; row++) {
        var _IC = fl[row][fl[0].indexOf(H_ICODE)];
        var dlItem = [];
        var dlDataLink = [];

        // push in the order needed for the display columns
        dlItem.push(fl[row][fl[0].indexOf(H_LINEID)]); // col 1
        dlItem.push(_IC); // col 2
        dlItem.push(fl[row][fl[0].indexOf(H_IDESC)]); // col 3

        if (isCore) {
            dlItem.push(searchCLbyIC(_IC, clMfg)); // col 4
            dlItem.push(searchCLbyIC(_IC, clPartNo)); // col 5
            dlItem.push(fl[row][fl[0].indexOf(H_BRANCHSUPPLIER)]); // col 6
            dlItem.push(
                'PR: ' + fl[row][fl[0].indexOf(H_PONUM)]
                + ' /PO: ' + fl[row][fl[0].indexOf(H_BRANCHPO)]
                ); // col 7

            // get the data link, if available
            var _dataLink = searchCLbyIC(_IC, clDataLink);
            dlDataLink.push('=HYPERLINK(\"' + _dataLink + '\",\"' + _IC + '\")');
            dlDataLinks.push(dlDataLink);

        } // end if isCore

        if (!isCore) {
            dlItem.push(''); // col 4, no manufacturer for WT50
            dlItem.push(''); // col 5, no part number for WT50
            dlItem.push(fl[row][fl[0].indexOf(H_SUPPLIER)]); // col 6
            dlItem.push(WT_PREFIX + ' ' + PROJ_NUMBER() + ' ' + fl[row][fl[0].indexOf(H_PONUM)]); // col 7
        } // end is not Core

        dlItem.push(fl[row][fl[0].indexOf(H_ACTDEL)]); // col 8
        dlItem.push(fl[row][fl[0].indexOf(H_QTY)]); // col 9
        dlItem.push(fl[row][fl[0].indexOf(H_QTYLEFT)]); // col 10
        dlItem.push(fl[row][fl[0].indexOf(H_PUOM)]); // col 11

        dl.push(dlItem);

    } // end for loop

    if (dl.length > 0) {
        // get a range equal in row length to the filtered list and write anyway to reinsert the header
        SH.getRange((grHeaderRow + 1), //starting row index
            1, // starting column index
            dl.length, // number of rows to import
            dl[0].length // number of columns to import
            ).setValues(dl); // set this range with the contents of the filtered list.

        if (isCore) {
            // write the data links into the first column, if the purchasing route isCore
            SH.getRange((grHeaderRow + 1), //starting row index
                2, // starting column index
                dlDataLinks.length, // number of rows to import
                1 // number of columns to import
                ).setFormulas(dlDataLinks); // set this range with the formula contents of the filtered list.
        } // end if isCore
    } // end if display lisy exists

    // auto resize selected columns
    /*  SH.autoResizeColumn(3);
      if (isCore){
        SH.autoResizeColumn(4);
        SH.autoResizeColumn(5);
      }
      SH.autoResizeColumn(6);
    */

} // end fn:getFilteredGRList


function filterGRList(list, ddValue, ddOption) {
    var subList = [];
    // for the provided List, find all items for this dropDown option
    for (var i = (1); i < list.length; i++) {
        // check that the status is "pending delivery" and items are left to be delivered
        if (list[i][list[0].indexOf(H_STATUS)].toString().substr(0, 1) > 3 // status 2 is "ready to send"
            && list[i][list[0].indexOf(H_STATUS)].toString().substr(0, 1) < 7 // status 6 is "part delivered"
            && list[i][list[0].indexOf(H_QTYLEFT)] > 0) {

            var li = list[i][list[0].indexOf(ddOption)];//.toString();//store the list item as a string
            if (li.indexOf(ddValue) > -1) { //check if this item's option field contains a string match to the search value
                subList.push(list[i]);
            } // end if

        } // end if "pending delivery" and items left to deliver
    } // end for loop
    subList.unshift(getML()[0]);// add in the title row at the beginning
    return subList;
} // end fn:filterGRList

function setGRDropDown(list, ddCell_A1, ddType) {
    var aOpt = []; // an array of the dropdown options
    // for the provided List, find all available options for this Type. Start at i=1 to remove header
    // the header is unshifted back to the top of the array at the end of fn:filterList
    for (var i = 1; i < list.length; i++) {
        var opt = list[i][list[0].indexOf(ddType)];
        if (aOpt.indexOf(opt) < 0) {
            aOpt.push(opt); // if this type isn't already in the dropdown option list, add it
        } // end if
    } // end for loop
    aOpt.sort(); // sort alphabetically

    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.requireValueInList(aOpt, true);
    SS.getSheetByName(GOODS_RECEIVING_SHEET).getRange(ddCell_A1).setDataValidation(dv.build());

} // end fn:setGRDropDown

function markGoodsReceived() {

    SS.toast("Collecting Materials and Delivery Data", "Goods Receiving 1 of 3", 600);

    var grSH = SS.getSheetByName(GOODS_RECEIVING_SHEET);
    var grSHlr = grSH.getLastRow();

    // goods recieving columns

    var grLineIDRange = grSH.getRange((grHeaderRow + 1), 1, grSHlr, 1);
    var grLineIDVals = grLineIDRange.getValues();

    var grQTYRemRange = grSH.getRange((grHeaderRow + 1), 10, grSHlr, 1);
    var grQTYRemVals = grQTYRemRange.getValues();

    var grQTYRcvdRange = grSH.getRange((grHeaderRow + 1), 12, grSHlr, 1);
    var grQTYRcvdVals = grQTYRcvdRange.getValues();

    var grDateRcvdRange = grSH.getRange((grHeaderRow + 1), 13, grSHlr, 1);
    var grDateRcvdVals = grDateRcvdRange.getValues();

    var grNotesRange = grSH.getRange((grHeaderRow + 1), 14, grSHlr, 1);
    var grNotesVals = grNotesRange.getValues();

    // materials sheet columns

    var mtLineIDRange = matSH.getRange(matFirst, col(H_LINEID).n, matLast, 1);
    var mtLineIDVals = mtLineIDRange.getValues();
    // .join().split(',') is used to convert the 2d arrays into 1d arrays of string
    mtLineIDVals = mtLineIDVals.join().split(',');

    var mtQTYRcvdRange = matSH.getRange(matFirst, col(H_QTYRCVD).n, matLast, 1);
    var mtQTYRcvdVals = mtQTYRcvdRange.getValues();

    var mtDateRcvdRange = matSH.getRange(matFirst, col(H_DATEDELIVERED).n, matLast, 1);
    var mtDateRcvdVals = mtDateRcvdRange.getValues();

    var mtNotesRange = matSH.getRange(matFirst, col(H_DELCOMMENTS).n, matLast, 1);
    var mtNotesVals = mtNotesRange.getValues();

    var mtStatusRange = matSH.getRange(matFirst, col(H_STATUS).n, matLast, 1);
    var mtStatusVals = mtStatusRange.getValues();

    SS.toast("Matching Delivery Data", "Goods Receiving 2 of 3", 600);

    for (var g = 0; g < grQTYRcvdVals.length; g++) {

        if (grQTYRcvdVals[g][0] > 0) { //check if recevied qty is greater than 0

            // get the index number from the materials sheet for this lineID
            var mtIndex = mtLineIDVals.indexOf(grLineIDVals[g][0]);
            Logger.log(mtIndex);
            mtQTYRcvdVals[mtIndex][0] = grQTYRcvdVals[g][0];
            mtDateRcvdVals[mtIndex][0] = grDateRcvdVals[g][0];
            mtNotesVals[mtIndex][0] = grNotesVals[g][0];

            if (grQTYRcvdVals[g][0] < grQTYRemVals[g][0]) { mtStatusVals[mtIndex][0] = '6 Part Delivered'; }
            if (grQTYRcvdVals[g][0] >= grQTYRemVals[g][0]) { mtStatusVals[mtIndex][0] = '7 Delivered'; }

        } // end if qty>0 goods received

    } // end for loop

    SS.toast("Writing Delivery Data", "Goods Receiving 3 of 3", 600);

    mtQTYRcvdRange.setValues(mtQTYRcvdVals);
    mtDateRcvdRange.setValues(mtDateRcvdVals);
    mtNotesRange.setValues(mtNotesVals);
    mtStatusRange.setValues(mtStatusVals);

    SS.toast("Done", "Goods Received", 5);

    grSH.getRange(ddMarkReceived).setValue('Mark Goods Received');

    grSH.getRange(ddPurchasingRoute).setValue('GO!');
    setGRFilters();

} // end fn:markGoodsReceived

/*
 * End Functions originally in GoodsReceiving.gs
 */

/*
 * Functions originally in Materials.gs
 */
var matSH = SS.getSheetByName(MATERIALS_SHEET);
var matFirst = MHRI + 2;
var matLast = matSH.getLastRow();

function matEdit(e) {

    /* write the supplier details into a note if the suppler dropdown was edited */
    if (e.range.getColumn() == col(H_SUPPLIER).n) {
        e.range.clearNote();
        matSH.getRange(col(H_ADMINCODE).a + e.range.getRow()).setValue('');
        var supplier = getSupplier(e.value);
        if (supplier.name == e.value) {
            var hoverNote = supplier.contact + '\n' + supplier.tel + '\n' + supplier.email;
            e.range.setNote(hoverNote);
            // set the WT Admin Code
            matSH.getRange(col(H_ADMINCODE).a + e.range.getRow()).setValue(supplier.admin);
        } // end if supplier has been found from lookup
    } // end if supplier column has been edited


    /* if the status has been set to a void value, add note to the PDN and PDC boxes then set PDN to void */
    if (e.range.getColumn() == col(H_STATUS).n) {
        var thisPDN = matSH.getRange(col(H_PDN).a + e.range.getRow());
        var thisPDC = matSH.getRange(col(H_PDC).a + e.range.getRow());

        if (e.value.substr(0, 1) >= VOID_PREFIX) {
            if (thisPDN.getNote() == '') {
                thisPDN.setNote('was: ' + thisPDN.getValue());
                thisPDC.setNote('was: ' + thisPDC.getValue());
                thisPDN.setValue('void');
            }
        } // if the status was set to void

        if (e.value.substr(0, 1) < VOID_PREFIX) {
            if (thisPDN.getNote() != '') {
                thisPDN.setValue(thisPDN.getNote().substr(5));
                thisPDN.clearNote();
                thisPDC.clearNote();
            }
        } // if the status was reset from void


    } // end if status column has been edited


} // end fn:matEdit

function matCleanup() {

    matDropDowns();
    matDateFormat();
    matNumberFormat();
    matSetFormulae();
    //  matSH.sort(col(H_PONUM).n);

} // end fn:matCleanup

function matDropDowns() {

    // get dropdowns from CentralData
    var CDD = getCentralDropDowns();

    // get suppliers from central and local
    var allSuppliers = getSupplier('*all*');
    var supplierNames = [];
    for (; allSuppliers.length > 0;) {
        supplierNames.unshift(allSuppliers.pop()[0]);
    }
    // remove the header from the central data list
    supplierNames.shift();

    var pdnDDSource = CDD.ddPDN; // set the Project Dimension dd as all PDNs
    // if there is a budget, limit the PDNs to only those with a budgeted value
    /*  if(SS.getRangeByName("BudgetTotal").getValue()>0){
        pdnDDSource = []; // clear out the source
        var budgetSubmitted = SS.getRangeByName("BudgetSubmitted").getValues();
        var budgetPDCs = SS.getRangeByName("BudgetPDCs").getValues();

        for (var i=0;i<budgetPDCs.length;i++){ // loop through pdc column on the budget sheet
          var pdcIndex = CDD.ddPDC.indexOf(budgetPDCs[i][0]); // check if the the budget pdc exists
          if (pdcIndex>-1  && budgetSubmitted[i][0] > 0){
            pdnDDSource.push(CDD.ddPDN[pdcIndex]);
          }// if this is a valid PDC and has a budgeted value
        }// end loop through budget PDCs
      }// if there is a budget value
      pdnDDSource.push(CDD.ddPDN.pop()); // add back in the last entry ("void") from the centrall dropdown
    */
    var columns = [ // an array of columnHeader and dropDown datasource pairs
        H_TYPE, CDD.ddType
        , H_TEAM, CDD.ddTeams
        , H_PUOM, CDD.ddUoM
        , H_BUOM, CDD.ddUoM
        , H_PDN, pdnDDSource
        , H_VATRATE, CDD.ddVATRates
        , H_SUPPLIER, supplierNames
        , H_EMERGENCY, ['Yes', 'No']
        , H_STATUS, CDD.ddStatusBPR
        , H_HIRESTATUS, CDD.ddStatusHire
    ];

    for (; columns.length > 0;) {
        var thisCol = columns.shift(); // the column header (first of the array pair)
        var colRange = col(thisCol).a + (MHRI + 2) + ':' + col(thisCol).a; // the column reference in the for A1:A
        var dv = SpreadsheetApp.newDataValidation();
        dv.setAllowInvalid(false);
        var _dd = columns.shift(); // the dropdown source datalist, (second of the array pair)
        dv.requireValueInList(_dd, true);
        SS.getSheetByName(MATERIALS_SHEET).getRange(colRange).setDataValidation(dv.build());
    } // end for loop

} // end fn:matDropDowns

function matDateFormat() {

    /* * * * *
     * column headers to format as date/time
     * * * * */
    var dtColumns = [
        col(H_ACTDEL).a
        , col(H_OFF).a
        , col(H_POCREATED).a
    ];

    for (; dtColumns.length > 0;) {
        var _col = dtColumns.pop();
        var fRange = matSH.getRange(_col + matFirst + ":" + _col + matLast);
        var fDT = []; // an array of number formats for datetime
        var fD = []; // an array of number formats for just the date

        for (var i = fRange.getNumRows(); i > 0; i--) {
            // push format inside of [] so that fDT and fD become a 2d object/array (rows x 1col)
            fDT.push(["dd/MM/yyyy HH:mm:ss"]);
            fD.push(["dd/MM/yyyy"]);
        }
        fRange.setNumberFormats(fDT);
        fRange.setNumberFormats(fD);
        fRange.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
    } // end for loop

} // end fn:matDateFormat

function matNumberFormat() {

    /* * * * *
     * column headers to format
     * * * * */
    var colNumFormat = [ // an array of columnHeader and formatString pairs
        col(H_ICODE).a, '@STRING@'
        , col(H_UNIT).a, '0.00'
        , col(H_NET).a, '0.00'
        , col(H_VATRATE).a, '0%'
        , col(H_VATVALUE).a, '0.00'
        , col(H_LINEVALUE).a, '0.00'
    ];

    for (; colNumFormat.length > 0;) {
        var _col = colNumFormat.shift();
        var _format = colNumFormat.shift();
        var fRange = matSH.getRange(_col + matFirst + ":" + _col + matLast);
        var fNum = []; // an array of number formats 2 dp number

        for (var i = fRange.getNumRows(); i > 0; i--) {
            // push format inside of [] so that fNum become a 2d object/array (rows x 1col)
            fNum.push([_format]);
        }
        fRange.setNumberFormats(fNum);
    } // end for loop

} // end fn:matNumberFormat

function matSetFormulae() {

    var formulaColumns = [
        col(H_NET).a
        , '=IF(AND(EQ(' + col(H_QTY).a + matFirst + ',""),EQ(' + col(H_UNIT).a + matFirst + ',"")),"",' + col(H_QTY).a + matFirst + '*' + col(H_UNIT).a + matFirst + ')'

        , col(H_VATVALUE).a
        , '=IF(EQ(' + col(H_VATRATE).a + matFirst + ',""),"",' + col(H_NET).a + matFirst + '*(RIGHT(' + col(H_VATRATE).a + matFirst + ',(LEN(' + col(H_VATRATE).a + matFirst + ')-2))))'

        , col(H_LINEVALUE).a
        , '=IF(AND(EQ(' + col(H_NET).a + matFirst + ',"")),"",' + col(H_NET).a + matFirst + '+' + col(H_VATVALUE).a + matFirst + ')'

        , col(H_PDC).a
        , '=IF(EQ(' + col(H_PDN).a + matFirst + ',""),"",VLOOKUP(' + col(H_PDN).a + matFirst + ',PD_Range,2,FALSE))'

        , col(H_QTYLEFT).a
        , '=IF(EQ(' + col(H_QTY).a + matFirst + ',""),"",' + col(H_QTY).a + matFirst + '-' + col(H_QTYRCVD).a + matFirst + ')'
    ];

    // set the initial row formulae, then copy to the rest of the column
    for (; formulaColumns.length > 0;) {
        var initCol = formulaColumns.shift();
        var initRange = matSH.getRange(initCol + matFirst);
        initRange.setFormula(formulaColumns.shift());
        initRange.copyTo(matSH.getRange(initCol + (matFirst + 1) + ":" + initCol + matLast));
    }

} // end fn:matSetFormulae

// function to return the column letter based on the header
function col(columnHeader) {
    // import header row, only if not previously imported during the scope of this script call
    if (!hRow) { hRow = matSH.getRange((MHRI + 1), 1, 1, matSH.getLastColumn()).getValues(); }
    var hIndex = hRow[0].indexOf(columnHeader);
    if (hIndex > -1) {
        var retVal = { i: hIndex, n: (hIndex + 1), a: alphaCols[hIndex] };
        return retVal;
    }
} // end fn:col

function updateStatustoSent() {

    UI = SpreadsheetApp.getUi();

    // thisRowIndex is the current row -1 to allow for zero index in arrays
    var thisRowIndex = SH.getActiveCell().getRow() - matFirst;

    // if a cell is selected in a header row, not an item row
    if (thisRowIndex <= MHRI) {
        var wrongRow = UI.alert("Status Update Error",
            "You have currently selected one of the header rows."
            + "\nPlease select a cell in a row from the materials list.",
            UI.ButtonSet.OK);
        return false;
    }

    // if not on the materials sheet
    if (SH.getName() != MATERIALS_SHEET) {
        var uiResponse = UI.alert("Status Update Error",
            "You are not on the " + MATERIALS_SHEET + " sheet.",
            UI.ButtonSet.OK);
        return false;
    }

    // get the PR numbers and Status Columns

    var statusBPRReady = '3 BPR Prepared';
    var statusBPRSent = '4 BPR Sent';

    var mtPRNoRange = matSH.getRange(matFirst, col(H_PONUM).n, matLast, 1);
    var mtPRNoVals = mtPRNoRange.getValues();


    var mtStatusRange = matSH.getRange(matFirst, col(H_STATUS).n, matLast, 1);
    var mtStatusVals = mtStatusRange.getValues();

    var thisBPRNo = mtPRNoVals[thisRowIndex][0];
    var okToUpdate = true; // a bit to check whether the update is possible

    for (var u = 0; u < mtPRNoVals.length; u++) {
        if (mtPRNoVals[u][0] == thisBPRNo) {
            if (mtStatusVals[u][0] != statusBPRReady) { okToUpdate = false; }
            if (mtStatusVals[u][0] == statusBPRReady) { mtStatusVals[u][0] = statusBPRSent; }
        } // end if this BPR is THIS BPR
        if (!okToUpdate) { break; }
    }// end loop

    if (!okToUpdate) {
        var uiResponse = UI.alert("Status Update Error",
            "Not all line items for BPR " + thisBPRNo + " have a status of \"" + statusBPRReady + "\".",
            UI.ButtonSet.OK);
        return false;
    }

    if (okToUpdate) {
        mtStatusRange.setValues(mtStatusVals);
    }


} // end fn:updateStatustoSent

/*
 * End Functions originally in Materials.gs
 */

/*
 * Functions originally in NonCore.gs
 */

function ncEdit(e) {
    var eRangeA1 = e.range.getA1Notation();

    if (eRangeA1 == ddProcessNonCore && e.value == 'GO!') {
        SS.toast("Please wait...", "Working", 600);
        processNonCore();
        e.range.setValue('Add to Materials Tracker');
    }
}


function processNonCore() {

    SS.toast("Processing Non Core Items", "Non Core 1 of 3", 600);

    var ncSH = SS.getSheetByName(NONCORE_SHEET);
    var ncRange = ncSH.getRange((ncHeaderRow + 1), 1, ncSH.getLastRow(), ncSH.getLastColumn());
    var ncValues = ncRange.getValues();

    var ncToAdd = [];

    // loop through the non-core items and format/concantenate accordingly

    for (var n = 0; n < ncValues.length; n++) {
        var _ncToAdd = [];
        var _ncIC = '';
        var _ncID = '';
        var _ncType = COREEXTRA;
        var _ncTemp = '';

        if (ncValues[n][0] != '') { // supplier catalog number
            _ncTemp = ncValues[n][0].toString().trim();
            _ncIC = '(' + _ncTemp + ')';
            _ncID = _ncIC;
        }

        if (ncValues[n][1] != '') { // manufacturers part number
            _ncTemp = ncValues[n][1].toString().trim().toUpperCase();
            if (_ncID != '') { _ncID += ' '; } // insert a trailing space if text already exists
            _ncID += _ncTemp;
        }

        if (ncValues[n][2] != '') { // manufacturers name
            _ncTemp = ncValues[n][2].toString().trim().toUpperCase();
            if (_ncID != '') { _ncID += ' '; } // insert a trailing space if text already exists
            _ncID += _ncTemp;
        }

        if (ncValues[n][3] != '') { // manufacturers brand name
            _ncTemp = ncValues[n][3].toString().trim().toUpperCase();
            if (_ncID != '') { _ncID += ' '; } // insert a trailing space if text already exists
            _ncID += _ncTemp;
        }

        if (ncValues[n][4] != '') { // first descriptive noun
            _ncTemp = ncValues[n][4].toString().trim().toUpperCase();
            if (_ncID != '') { _ncID += ' '; } // insert a trailing space if text already exists
            _ncID += _ncTemp + ':';
        }

        if (ncValues[n][5] != '') { // adjectives
            _ncTemp = ncValues[n][5].toString().trim().toLowerCase();
            if (_ncID != '') { _ncID += ' '; } // insert a trailing space if text already exists
            _ncID += _ncTemp;
        }

        // check that there is SOMETHING in the description
        if (_ncID != '') {
            _ncToAdd.push(_ncID);
            _ncToAdd.push(''); // blank line for the item code (was _ncIC instead of '')
            _ncToAdd.push(_ncType);
            ncToAdd.push(_ncToAdd);
        }

    } // end for loop

    // get a range on the main materials tracker equal to the non core items to add and write them in
    // this code is copied from the CoreList fn:sendBasket

    SS.toast("Adding Non Core Items", "Non Core 2 of 3", 600);

    matSH = SS.getSheetByName(MATERIALS_SHEET)
  var curMatRange = matSH.getRange(1, //starting row index
        2, // starting column index
        matSH.getLastRow(), // number of rows to import
        matSH.getLastColumn() // number of columns to import
        );

    var curMat = curMatRange.getValues();
    var blankRow = 0;

    for (var i = curMat.length - 1; i > MHRI; i--) {
        // check for the first blank row
        if (
            (curMat[i][curMat[MHRI].indexOf(H_IDESC)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_ICODE)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_TYPE)] != '')
            ) { blankRow = i + 2; break; }
    } // end for loop

    matSH.getRange(blankRow, //starting row index
        col(H_IDESC).n, // starting column index
        ncToAdd.length, // number of rows to import
        ncToAdd[0].length // number of columns to import
        ).setValues(ncToAdd);

    SS.toast("Clearing Non Core (Free Text)", "Non Core 3 of 3", 5);

    // once complete, clear the existing non core entereed items
    ncRange.clear();
}

/*
 * End Functions originally in NonCore.gs
 */

/*
 * Functions originally in PO 3 and 4.gs
 * POs3_createPO
 * nextPOPRNumber
 * getSupplier
 * POs4_updateMaterialsList
 * generatePO
 */

function POs3_createPO() {

    var appNameCell = 'H3';
    var appPONumberCell = 'H4';
    var appSubTotalCell = 'L33';

    var poMain = 'Main Purchase Order';
    var poMain_StartRow = 20;
    var poMain_EndRow = 29;
    var poAppendix = 'Appendix';
    var poAppendix_StartRow = 7;
    var poAppendix_EndRow = 31;

    var poShippingInfo = SS.getRangeByName('projAddress').getValues();

    var newPONumber = nextPOPRNumber();
    var newPOName = WT_PREFIX + ' ' + PROJ_NUMBER + ' ' + newPONumber;

    var newPOFile = DriveApp.getFileById(PO_TEMPLATE_ID)
        .makeCopy(newPOName + ' ~ ' + poSupplierName,
        DriveApp.getFolderById(PO_FOLDER_ID));
    newPOFile.setOwner(DRIVE_OWNER);

    var newPOUrl = newPOFile.getUrl();

    var newPOID = newPOUrl.substring(newPOUrl.indexOf('/d/') + 3, newPOUrl.indexOf('/edit'));

    var retVal = { Num: newPONumber, Url: newPOUrl, ID: newPOID, Name: newPOName };

    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(newPOID));
    var PO = SpreadsheetApp.getActiveSpreadsheet();
    var POsh = SpreadsheetApp.getActiveSheet();

    var i = iIndices.length;
    var appSubs = []; // an array of appendix subtotals

    // if i>10, then calculate how many appendicies, a, are needed
    if (i > 10 && i <= 250) {
        var a = Math.floor(((i - 10) / 25) + 1); // because i = 25a + 10-a, that is 25 items on each appendix and 10-a number of items on the main PO to allow for an appendix line.

        var appNames = [];

        // create an appendix for each appendix needed
        for (var j = 0; j < a; j++) {
            POsh = PO.setActiveSheet(PO.getSheetByName(poAppendix)).copyTo(PO).setName(poAppendix + ' ' + alphaCols[j]);
            POsh.getRange(appPONumberCell).setValue(newPOName);
            POsh.getRange(appNameCell).setValue(POsh.getSheetName());
            appNames.push(POsh.getSheetName());
        }

        var appRow = ((i - 10 + a) % 25); // the remainder which is how many rows go on the last appendix

        // fill up all the appendices
        for (; appNames.length > 0;) {
            POsh = PO.getSheetByName(appNames.pop()).activate();
            var appendixRange = POsh.getRange(poAppendix_StartRow - 1, 1, appRow + 1, 8); // use poAppendix_StartRow-1 so that a 0 index array is the header and so appRow+1 for the range
            var appendixValues = appendixRange.getValues();
            var appendixPDNRange = POsh.getRange(poAppendix_StartRow - 1, 13, appRow + 1); // just one column
            var appendixPDNValues = appendixPDNRange.getValues();

            for (; appRow > 0; appRow--) {
                appendixValues[appRow][0] = (10 * (appRow));
                appendixValues[appRow][1] = iQty.pop();
                appendixValues[appRow][3] = iUoM.pop();
                appendixValues[appRow][4] = iDesc.pop();
                appendixValues[appRow][7] = iUnit.pop();
                appendixPDNValues[appRow][0] = iPDN.pop();
            }
            appendixRange.setValues(appendixValues);
            appendixPDNRange.setValues(appendixPDNValues);
            appSubs.push(POsh.getRange(appSubTotalCell).getValue());
            appRow = poAppendix_EndRow - poAppendix_StartRow + 1;
        }

    } // end if i>10

    // focus on the main PO page
    POsh = PO.setActiveSheet(PO.getSheetByName(poMain));

    // write the single values to the new PO
    PO.getRangeByName('wtPO').setValue(newPOName);
    PO.getRangeByName('sup_name').setValue(poSupplierName);
    var supplier = getSupplier(poSupplierName);
    if (supplier.name == poSupplierName) {
        PO.getRangeByName('sup_adr1').setValue('c/o ' + supplier.contact);
        PO.getRangeByName('sup_adr2').setValue(supplier.adr);
        PO.getRangeByName('sup_adr3').setValue(supplier.tel);
        PO.getRangeByName('sup_email').setValue(supplier.email);
        PO.getRangeByName('pay_terms').setValue(supplier.terms);
        PO.getRangeByName('cust_acct').setValue(supplier.account);
    }
    PO.getRangeByName('del_onby').setValue(delONBY);
    PO.getRangeByName('del_date').setValue(new Date(poDelivery));
    PO.getRangeByName('order_date').setValue(new Date());
    PO.getRangeByName('buyer_email').setValue(Session.getActiveUser().getEmail());

    // write the shipping information
    PO.getRangeByName('ship_adr1').setValue(poShippingInfo[0][0]);
    PO.getRangeByName('ship_adr2').setValue(poShippingInfo[1][0]);
    PO.getRangeByName('ship_adr3').setValue(poShippingInfo[2][0]);
    PO.getRangeByName('ship_adr4').setValue(poShippingInfo[3][0]);
    PO.getRangeByName('ship_adr5').setValue(poShippingInfo[4][0]);

    // get the body area of the main po page
    var mainPORange = POsh.getRange(poMain_StartRow - 1, 1, poMain_EndRow - poMain_StartRow + 2, 8);
    var mainPOValues = mainPORange.getValues();
    var mainPOPDNRange = POsh.getRange(poMain_StartRow - 1, 13, poMain_EndRow - poMain_StartRow + 2);
    var mainPOPDNValues = mainPOPDNRange.getValues();

    // fill the main PO page with line items
    for (; iQty.length > 0;) {
        mainPOValues[iQty.length][0] = (10 * (iQty.length));
        mainPOValues[iQty.length][1] = iQty.pop();
        mainPOValues[iUoM.length][3] = iUoM.pop();
        mainPOValues[iDesc.length][4] = iDesc.pop();
        mainPOValues[iUnit.length][7] = iUnit.pop();
        mainPOPDNValues[iPDN.length][0] = iPDN.pop();
    }

    // add any appendices references to the main page

    var r = poMain_EndRow - poMain_StartRow - appSubs.length + 2;
    var _appNo = appSubs.length;

    for (; appSubs.length > 0; r++) {
        mainPOValues[r][0] = (10 * r);
        mainPOValues[r][1] = 1;
        mainPOValues[r][3] = 'ea';
        mainPOValues[r][4] = '* * * SEE APPENDIX ' + alphaCols[_appNo - appSubs.length] + ' * * *';
        mainPOValues[r][7] = appSubs.pop();
        mainPOPDNValues[r][0] = '< < APPENDIX';
    }

    mainPORange.setValues(mainPOValues);
    mainPOPDNRange.setValues(mainPOPDNValues);

    // delete the template appendix  
    PO.deleteSheet(PO.getSheetByName(poAppendix));

    // return to original sheet
    SpreadsheetApp.setActiveSpreadsheet(SS);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SH).setActiveSelection(thisCell);

    // return the details of the new PO/PR to the main script
    return retVal;

} // enf fn:POs3_createPO

function nextPOPRNumber() {

    // if a PO number already exists (i.e. regeneration) just return the same PO number
    if (poNumber != '') { return poNumber; }

    // place holder for the next 4-digit PO number and the whole list of current PO.
    var nextNumber = '';
    var upperLimit = 2000;
    var lowerLimit = 0;
    var currentNumbers = [];

    /*  // set the upper limit of the number range for the PO/PR
      if (TYPE==CORELIST||COREEXTRA){lowerLimit=0;upperLimit=999;}
      if (TYPE==MATERIALS){lowerLimit=1000;upperLimit=1999;}
      if (TYPE==HIRE){lowerLimit=2000;upperLimit=2999;}
    */
    var POPRnumbers = SH.getRange((MHRI + 2), (wholeList[MHRI].indexOf(H_PONUM) + 1), SH.getLastRow(), 1).getValues();

    for (var i = 0; i < POPRnumbers.length; i++) {
        var n = POPRnumbers[i][0];
        if (lowerLimit < n && n < upperLimit) { currentNumbers.push(n); }
    }

    if (currentNumbers.length == 0) { currentNumbers.push(lowerLimit) }

    // sort the list of current PO numbers to get the highest one
    currentNumbers.sort(function (a, b) {return b - a });
    var _nextNumber = currentNumbers[0];

    _nextNumber = new Number(_nextNumber) + 1;
    nextNumber = _nextNumber.toString();
    nextNumber = '0000' + nextNumber;
    nextNumber = nextNumber.substr(-4);

    return nextNumber;

}

function getSupplier(supplierName): any
{
    // set up placeholders
    aCD_WTSuppliers = getWTSup();

    var _sheet = SS.getSheetByName(LOCAL_SUPPLIERS_SHEET);
    var localSuppliers = _sheet.getRange(4, //starting row index
        1, // starting column index                       
        (_sheet.getLastRow() - 3), // number of rows to import                           
        _sheet.getLastColumn() // number of columns to import
        ).getValues(); // an array of rows, each an array of columns wholeList[r-1][c-1]

    var allSuppliers = aCD_WTSuppliers.concat(localSuppliers);


    if (supplierName == '*all*') { return allSuppliers; }

    // the Index for the "found" supplier
    var I = -1;

    // search for the provided supplier by their name
    for (var i = 0; i < allSuppliers.length; i++) {
        if (allSuppliers[i][0] == supplierName) {
            I = i; // store the index in I
        }
    }

    var poSupplier = {
        name: '', contact: '', adr: '',
        tel: '', email: '',
        account: '', terms: '', admin: ''
    };

    if (I > -1) {
        poSupplier.name = supplierName;
        poSupplier.contact = allSuppliers[I][4]; // the WT Supplier Contact
        poSupplier.adr = ''; // leave this blank
        poSupplier.tel = allSuppliers[I][5]; // the WT Supplier telephone number
        poSupplier.email = allSuppliers[I][6]; // the WT Supplier email address
        poSupplier.account = allSuppliers[I][2]; // the WT Supplier WT Account number
        poSupplier.terms = allSuppliers[I][3]; // the WT Supplier payment terms
        poSupplier.admin = allSuppliers[I][1]; // the WT Supplier admin code
    }

    return poSupplier;

} // end fn:getWTSupplier

/* * * * * * *
 * STAGE 4 of PO Generation Materials - write the PO/PR number
 * 
 * * * * * * */

function POs4_updateMaterialsList(POs3) {

    // import the three columns that need updating
    var statusCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_STATUS) + 1), SH.getLastRow(), 1);
    var poNumCol = SH.getRange((MHRI + 2), (wholeList[MHRI].indexOf(H_PONUM) + 1), SH.getLastRow(), 1);
    var poCreatedCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_POCREATED) + 1), SH.getLastRow(), 1);

    var statusValues = statusCol.getValues();
    var poNumFormulas = poNumCol.getFormulas();
    var poCreatedValues = poCreatedCol.getValues();

    var poprLink = '=HYPERLINK(\"' + POs3.Url + '\",\"' + POs3.Num + '\")';

    SS.toast("PO created, adding links to materials list", "Step 5 of 5", 60);

    // remove the references for the old file for regenerated items
    while (aRegenIndices.length) {
        var r = aRegenIndices.pop();
        statusValues[r][0] = '';
        poNumFormulas[(r - MHRI - 1)][0] = '';
        poCreatedValues[r][0] = '';
    }

    var _status = null;
    var _statusArray = getCentralDropDowns().ddStatusWT50;
    for (; _statusArray.length > 0;) {
        // get the status dropdown value whose prefix matches the regen prefix
        if (_statusArray[0].substr(0, 1) == REGEN_PREFIX) { _status = _statusArray[0]; }
        _statusArray.shift();
    }

    // enter the values for the newly created file into the correct rows.
    while (iIndices.length) {
        var i = iIndices.pop();
        statusValues[i][0] = _status;
        poNumFormulas[(i - MHRI - 1)][0] = poprLink;
        poCreatedValues[i][0] = new Date();
    }

    statusCol.setValues(statusValues);
    poNumCol.setFormulas(poNumFormulas);
    poCreatedCol.setValues(poCreatedValues);

    return true;
}

/*
 * End Functions originally in PO 3 and 4.gs
 */

/*
 * Functions originally in POPR 0.gs
 */


/*
 * ciriteria for a PO:
 * same... supplier, delivery date and type (possibly limit to same department)
 *
*/
function generatePO(formObject) {

    sameDept = formObject.sameDept;
    delONBY = formObject.delivery;

    UI = SpreadsheetApp.getUi();
    var SHname = SH.getName();
    thisCell = SH.getActiveCell();

    if (SHname == MATERIALS_SHEET) {

        // thisRowIndex is the current row -1 to allow for zero index in arrays
        var thisRowIndex = SH.getActiveCell().getRow() - 1;

        // if a cell is selected in a header row, not an item row
        if (thisRowIndex <= MHRI) {
            var wrongRow = UI.alert("Generate PO Error",
                "You have currently selected one of the header rows."
                + "\nPlease select a cell in a row from the materials list.",
                UI.ButtonSet.OK);
            return false;
        }

        // * * * * stage 1, get the materials * * * 
        var POs1 = POs1_getItems(thisRowIndex); // returns an array of row numbers of any incomplete items

        // * * * * stage 2, handle any incomplete items * * *
        var POs2 = POs2_handleIncompleteItems(POs1); // returns true or false to proceed

        // if there are more than 250 items and a PO is needed
        if (iIndices.length > 250) {
            var tooManyItems = UI.alert("Too Many Items for One PO",
                "The maximum number of line items for a purchase order is 250.",// The first 250 valid items will be added to this PO then the remainder will be added to a new PO"
                UI.ButtonSet.OK);
        }
        // ********* this needs doing    

        /*    // if there are less than 250 items and a PO is needed
            if (POs2 && iIndices.length <= 250 && TYPE!=CORELIST && TYPE!=COREEXTRA) {
              var POs2_msg = (iIndices.length == 1) ? "Creating PO with 1 valid item.":"Creating PO with "+iIndices.length+" valid items.";
              SS.toast(POs2_msg,"Step 4 of 5",60);
              var POs3 = POs3_createPO();
              // clean up and add the PO links
              var POs4 = POs4_updateMaterialsList(POs3);
              if (POs4){SS.toast("PO "+POs3.Num+" has been successfully created","Done!",5);}
            }
        */
        // if a BPR is needed
        if (POs2) {

            // setup holders for the completed message

            // an array to hold all the produced BPR data
            var BPRs = [];

            // use a loop to create the PRs
            for (var noOfPRs = Math.ceil(iIndices.length / PR_ITEMS_LIMIT); noOfPRs > 0; noOfPRs--) {

                var _msgPR = (noOfPRs == 1) ? "1 PR with" : noOfPRs + " PRs from";
                var msgPR = (iIndices.length == 1) ? "Creating 1 PR with 1 valid item." : "Creating " + _msgPR + " " + iIndices.length + " valid items.";
                SS.toast(msgPR, "Step 4 of 5", 60);

                var PRs3 = PRs3_createPR();
                BPRs.push(PRs3);
                // clean up and add the PO links
                var PRs4 = PRs4_updateMaterialsList(PRs3);

            } // end for loop
            SS.toast('Thank you for waiting', 'Done', 5);

            /*      // ask if newly created BPRs should be emailed now
                  var bprTitle = '';
                  var bprMessage = '';
  
                  var noBPRs = BPRs.length;
      
                  if (noBPRs==1) { bprTitle = noBPRs+" Purchase Request Created"; }
                  if (noBPRs==1) { bprMessage = noBPRs+" has been successfully created. Would you like to email it to "+poSupplier.name+" now?"; }
      
                  if (noBPRs>1) { bprTitle = noBPRs+" Purchase Requests Created"; }
                  if (noBPRs>1) { bprMessage = noBPRs+" have been successfully created. Would you like to email them to "+poSupplier.name+" now?"; }
      

                  if (noBPRs>0) {

                    var bprEmail = UI.alert(bprTitle, bprMessage, UI.ButtonSet.YES_NO);

                    // Process the user's response.
                    if (bprEmail == UI.Button.YES) {
                      var bprEmailToast = (noBPRs == 1) ? "Sending "+noBPRs+" PR to "+poSupplier.name:"Sending "+noBPRs+" PRs to "+poSupplier.name;
                      SS.toast(bprEmailToast,"Emailing PR",60);
          
                      // call the PR_Email function and wait for a true return
                      if (PR_Email(BPRs)){
                        var bprEmailedToast = (noBPRs == 1) ? noBPRs+" PR sent to "+poSupplier.name:noBPRs+" PRs sent to "+poSupplier.name;
                        SS.toast(bprEmailedToast,"Done!",5);
                      }
          
                    } // end if user said yes to sending email(s)
                  } // end if BPRs were created
            */
        } // end if BPR is needed

    } // end if correct sheet

    // if not on the materials sheet
    if (SHname != MATERIALS_SHEET) {
        SS.getSheetByName(MATERIALS_SHEET).activate();
        var uiResponse = UI.alert("Generate PO/PR",
            ". . . Changing Sheet to Materials List\nPlease click in a row you would like to generate a PO/PR for, then choose \"Generate PO/PR\" again.",
            UI.ButtonSet.OK);
    } // end if wrong sheet

} // end fn:generatePO_materials

/*
 * End Functions originally in POPR 0.gs
 */

/*
 * Functions originally in POPR 1.gs
 * 
 */

/* * * * * * *
* STAGE 1 of PO Generation Materials - get materials and validate items
*   get the whole sheet of items
*   extract only matching Supplier and Delivery date
*   Alert if any items are blank or else
*     write item arrays for ready for PO
* * * * * * */

function POs1_getItems(thisRowIndex) {

    SS.toast("Compiling PO/PR Items, please wait", "Step 1 of 5", 5);

    // import materials list
    wholeRange = SH.getRange(1, //starting row index
        1, // starting column index
        SH.getLastRow(), // number of rows to import
        SH.getLastColumn() // number of columns to import
        );

    wholeList = wholeRange.getValues(); // an array of rows, each an array of columns wholeList[r-1][c-1]

    var incompleteItems = []; // array for incomplete items, not added to the PO

    // single values for PO entries from materials list of thisRow
    poSupplierName = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_SUPPLIER)];
    TYPE = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_TYPE)];

    poDelivery = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_ACTDEL)];
    var _poDelivery = '';
    if (poDelivery != '') {
        _poDelivery = new Date(poDelivery);
        // the date is set to 8am to remove a problem of GMT dates being shown as the day before when the "now" date is BST
        var _poDelivery = Utilities.formatDate(new Date(poDelivery.setHours(8)), "GMT", "E, dd-MMM-yyyy");
    }
    poNumber = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_PONUM)];
    requestingDept = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_TEAM)];

    var thisLineStatus = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_STATUS)];

    branchPONumber = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_BRANCHPO)];

    if (wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_EMERGENCY)] == 'Yes') { isEmergency = true; }

    // if the PO Generation has been confirmed based on this line item, collect other matching items
    if (confirmGeneratePO(poSupplierName, _poDelivery, poNumber, requestingDept, thisLineStatus, isEmergency, branchPONumber)) {

        SS.toast("Sorting PO Items, please wait", "Step 3 of 5", 5);

        // loop through wholeList to extract the correct items
        for (var i = MHRI + 1; i < wholeList.length; i++) {
            var _tempDelDate = new Date(wholeList[i][wholeList[MHRI].indexOf(H_ACTDEL)]);
            var tempDelDate = Utilities.formatDate(new Date(_tempDelDate.setHours(8)), "GMT", "E, dd-MMM-yyyy");

            var _thisBranchPO = wholeList[i][wholeList[MHRI].indexOf(H_BRANCHPO)];

            if (
                // either if the PO number is blank OR the poNumber is this lines PO number AND is allowed to regenerate
                ((wholeList[i][wholeList[MHRI].indexOf(H_PONUM)] == '' && wholeList[i][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) == ITEMREADY_PREFIX)
                || (wholeList[i][wholeList[MHRI].indexOf(H_PONUM)] == poNumber && wholeList[i][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) == REGEN_PREFIX))

                // either if items don't need to be from the same department OR if they do, this item was requested by the same department as the original item
                && (!sameDept || (sameDept && wholeList[i][wholeList[MHRI].indexOf(H_TEAM)] == requestingDept))
                // the supplier is the same
                && wholeList[i][wholeList[MHRI].indexOf(H_SUPPLIER)] == poSupplierName
                // the type is the same
                && wholeList[i][wholeList[MHRI].indexOf(H_TYPE)] == TYPE
                // the delivery date is the same
                && ((!isEmergency && tempDelDate == _poDelivery) || (isEmergency))
                // if this IS an emergency order, check that the branch PO number is the same
                && (!isEmergency || (isEmergency && _thisBranchPO == branchPONumber))
                ) {

                var _isEmergency = false;
                if (wholeList[i][wholeList[MHRI].indexOf(H_EMERGENCY)] == 'Yes') { _isEmergency = true; }

                var _Qty = new Number(wholeList[i][wholeList[MHRI].indexOf(H_QTY)]);
                var _UoM = wholeList[i][wholeList[MHRI].indexOf(H_PUOM)];
                var _Factor = wholeList[i][wholeList[MHRI].indexOf(H_FACTOR)];
                var _BUoM = wholeList[i][wholeList[MHRI].indexOf(H_BUOM)];
                var _DESC = wholeList[i][wholeList[MHRI].indexOf(H_IDESC)];
                var _CODE = wholeList[i][wholeList[MHRI].indexOf(H_ICODE)];
                //        if (_CODE!='' && TYPE!=CORELIST && TYPE!=COREEXTRA) {_DESC += ' ['+_CODE+']';} // quick append of code to description if this is a PO, not PR
                var _UNIT = new Number(wholeList[i][wholeList[MHRI].indexOf(H_UNIT)]);
                var _PDN = wholeList[i][wholeList[MHRI].indexOf(H_PDN)];
                var _PDC = wholeList[i][wholeList[MHRI].indexOf(H_PDC)];
                var _NOTEtemp = wholeList[i][wholeList[MHRI].indexOf(H_NOTES)];
                var _NOTE = '';
                if (delONBY == 'ON:' && !isEmergency) { _NOTE = 'Please deliver ON: ' + _poDelivery; }
                if (delONBY == 'BY:' && !isEmergency) { _NOTE = 'Please deliver from 2 days before: ' + _poDelivery; }
                if (isEmergency) { _NOTE = 'Emergency Order: ' + branchPONumber; }
                if (_NOTE != '' && _NOTEtemp != '') { _NOTE += ' ~ ' + _NOTEtemp } // append actual note to delivery request if actual note is present
                if (_NOTE == '' && _NOTEtemp != '') { _NOTE += _NOTEtemp } // append actual note to delivery request if actual note is present

                // if this is a PR for a hire item, get the off hire date
                if (TYPE == HIRE) {
                    var offHire = wholeList[i][wholeList[MHRI].indexOf(H_OFF)];

                    if (offHire != '') {
                        var _offHire = new Date(offHire);
                        // the date is set to 8am to remove a problem of GMT dates being shown as the day before when the "now" date is BST
                        offHire = Utilities.formatDate(new Date(_offHire.setHours(8)), "GMT", "E, dd-MMM-yyyy");
                        _NOTE += '{Hire from: ' + tempDelDate + ' to: ' + offHire + '}'; // update the description with item specific on hire / off hire details.
                    } // end if valid off hire date

                    if (offHire == '' || _offHire < _tempDelDate) {  // if the off hire date is blank or is before the on hire / delivery date
                        incompleteItems.push(i + 1); // ,register an incomplete item
                        continue; // and continue to the next loop
                    }
                } // end if a HIRE PR

                // if all elements are present, push into array
                if (_Qty > 0 && _UoM != '' && _Factor != '' && _BUoM != '' && _DESC != '' && (_CODE != '' || TYPE != CORELIST) && _UNIT > 0 && _PDN != '' && _PDC != ''
                    && ((_isEmergency && _thisBranchPO == branchPONumber) || (!_isEmergency && _thisBranchPO == ''))
                    && ((_isEmergency && tempDelDate == _poDelivery) || (!_isEmergency))
                    ) {
                    iQty.push(_Qty);
                    iUoM.push(_UoM);
                    iFactor.push(_Factor);
                    iBUoM.push(_BUoM);
                    if (TYPE == CORELIST || TYPE == COREEXTRA) { iCode.push(_CODE); } // separately include the item code if a PR is being created rather than a PO
                    iDesc.push(_DESC);
                    iUnit.push(_UNIT);
                    iPDN.push(_PDN);
                    iPDC.push(_PDC);
                    iNote.push(_NOTE);
                    iIndices.push(i);
                } // end if elements all are present for this item        
                else { incompleteItems.push(i + 1); } // +1 to equate to row number instead of index number

            } // end if po item matches original
        } // end for loop

        return incompleteItems;

    } // end if confirmGeneratePO(...)

    return null;

} // end fn:POs1_getMaterialsValidate

function confirmGeneratePO(poSupplier, poDelivery, poNumber, requestingDept, thisLineStatus, isEmergency, branchPONumber) {

    POorPR = 'PR'; OrderOrRequest = 'Request';

    // if a PO number is already present
    if (poNumber != '') {

        var poRegeneratePossible = true;
        var alreadyPOmsg = "This item has already been added to a Purchase " + OrderOrRequest + ", " + poNumber + ".";
        var alreadyPObtn = UI.ButtonSet.YES_NO;
        var poLink = SH.getRange(
            thisCell.getRow(),
            (wholeList[MHRI].indexOf(H_PONUM) + 1)
            ).getFormula();
        var poID = poLink.substring(poLink.indexOf('/d/') + 3, poLink.indexOf('/edit'));

        // loop through wholeList to check if PO regeneration is possible (every line item that has that PO number must allow regeneration)
        for (var j = MHRI + 1; j < wholeList.length; j++) {
            if (wholeList[j][wholeList[MHRI].indexOf(H_PONUM)] == poNumber) {
                aRegenIndices.push(j + 1); // write the indices for the all currently assigned POs
                if (wholeList[j][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) != REGEN_PREFIX) {
                    poRegeneratePossible = false;
                } // end if PO can't be regenerated
            } // end if PO item match is found
        } // end whole list loop

        if (!poRegeneratePossible) {
            alreadyPOmsg += "\nPlease choose an item not yet ordered, as this " + POorPR + " cannot be regenerated.";
            alreadyPObtn = UI.ButtonSet.OK;
        }
        if (poRegeneratePossible) {
            alreadyPOmsg += "\nHowever, it is possible to regenerate this "
            + POorPR + " to include all valid items. Do you want to regenerate this " + POorPR + "?";
        }

        var alreadyPO = UI.alert("Generate " + POorPR + " Error", alreadyPOmsg, alreadyPObtn);

        if (!poRegeneratePossible || (alreadyPO == UI.Button.NO)) {
            SS.toast(POorPR + " generation cancelled", POorPR + " Cancelled", 5);
            return false;
        }

        if (alreadyPO == UI.Button.YES) {
            SS.toast("Deleting existing " + POorPR + " and removing links, please wait", "Deleting " + POorPR + " " + poNumber, 60);
            // mark for deletion
            var oldPO = DriveApp.getFileById(poID)
      oldPO.setName('TO DELETE >>> ' + oldPO.getName());

        } // end if YES to regenerate PO

    } // end if PO number already exists


    // if this is an emergency order, but no Branch PO number is present
    if (isEmergency && branchPONumber == '') {
        var noBranchPO = UI.alert("Generate " + POorPR + " Error",
            "This line item is set as part of an emergency order, but no Branch PO has been entered.",
            UI.ButtonSet.OK);
        // ***** possibly setActiveSelection to the cell that needs to be completed

        return false;
    }

    // if a Branch PO number is present, but this item is not set to be an emergency order
    if (branchPONumber != '' && !isEmergency) {
        var branchPObutNotEmergency = UI.alert("Generate " + POorPR + " Error",
            "This line item has a Branch PO entered, but is not set as part of an emergency order.",
            UI.ButtonSet.OK);
        // ***** possibly setActiveSelection to the cell that needs to be completed

        return false;
    }

    // if the supplier is empty
    if (poSupplier == '') {
        var noSupplier = UI.alert("Generate " + POorPR + " Error",
            "No Supplier has been set for this line item. Please choose a valid supplier.",
            UI.ButtonSet.OK);
        // ***** possibly setActiveSelection to the cell that needs to be completed

        return false;
    }

    // if the delivery date is empty
    if (poDelivery == '') {
        var noDeliveryDate = UI.alert("Generate " + POorPR + " Error",
            "No Delivery Date has been set for this line item. Please enter a valid date.",
            UI.ButtonSet.OK);
        // ***** possibly setActiveSelection to the cell that needs to be completed
        return false;
    }

    // if the status is not valid
    if (thisLineStatus.substr(0, 1) != ITEMREADY_PREFIX && !poRegeneratePossible) {
        var wrongStatus = UI.alert("Generate " + POorPR + " Error",
            "This line item does not have the correct status to be processed.",
            UI.ButtonSet.OK);
        // ***** possibly setActiveSelection to the cell that needs to be completed
        return false;
    }

    // if there is a Suppiler and Delivery date entered for the item in this row
    if ((poSupplier != '' || poRegeneratePossible) && poDelivery != '') {

        var msg = "Create a new " + POorPR + " for all to-be-ordered items from " + poSupplier + ", to be delivered " + delONBY + " " + poDelivery + "?";

        if (TYPE == HIRE) {
            msg = "Create a new " + POorPR + " for all to-be-hired items from " + poSupplier + ", to be delivered ON: " + poDelivery + "?"
            + "\n(Please note, hire " + POorPR + "s are always set to an ON delivery date, rather than BY)";
            delONBY = "ON:";
        }

        if (sameDept) { msg += "\n\n(Only items requested by \"" + requestingDept + "\" will be added to this " + POorPR + ")"; }

        if (isEmergency) {
            msg = "Create a new " + POorPR + " for all emergency ordered items from " + poSupplier + ", on " + branchPONumber + "?";
            sameDept = false;
        }

        var confirmPO = UI.alert("Generate " + POorPR + " for " + poSupplier, msg, UI.ButtonSet.YES_NO);
        // Process the user's response.
        if (confirmPO == UI.Button.YES) {
            SS.toast("Collecting " + POorPR + " Items, please wait", "Step 2 of 5", 5);
            return true;
        }
        if (confirmPO == UI.Button.NO) {
            SS.toast(POorPR + " generation cancelled by user", POorPR + " Cancelled", 5);
            return false;
        }
    }


} // end fn:confirmGeneratePO

/*
 * End Functions originally in POPR 1.gs
 */

/*
 * Functions originally in POPR 2.gs
 */

/* * * * * * *
* STAGE 2 of PO Generation Materials - handle any incomplete items
*   Alert if any items have missing information
*     and ask whether to proceed anyway or cancel and correct
* * * * * * */

function POs2_handleIncompleteItems(POs1) {

    // if there are no incomplete items and there is at least one valid item, return true
    if (POs1.length == 0 && iIndices.length > 0) { return true; }

    // if there are only incomplete items but no valid ones, alert and then return false
    if (POs1.length > 0 && iIndices.length == 0) {
        var nothingValidMessage = (POs1.length == 1) ? "Please check row " + POs1 + ", as it has missing information." : "Please check rows " + POs1 + ", as they have missing information.";
        var nothingValid = UI.alert("No Valid Items Found", nothingValidMessage, UI.ButtonSet.OK);
        var toastMsg1 = (POs1.length == 1) ? "Row to check: " + POs1 : "Rows to check: " + POs1;
        SS.toast(toastMsg1, "Incomplete Items", 60);
        return false;
    }

    // if there are any incomplete items and this is an emergency order, alert and then return false
    if (POs1.length > 0 && isEmergency) {
        var emergencyMessage = (POs1.length == 1) ? "Please check row " + POs1 + ", as it has missing information." : "Please check rows " + POs1 + ", as they have missing information.";
        var emergencyAlert = UI.alert("Emergency Order Incomplete", emergencyMessage, UI.ButtonSet.OK);
        var toastMsg1e = (POs1.length == 1) ? "Row to check: " + POs1 : "Rows to check: " + POs1;
        SS.toast(toastMsg1e, "Incomplete Items", 60);
        return false;
    }

    var incompleteItemsTitle = '';
    var incompleteItemsMessage = '';

    if (POs1.length > 1) { incompleteItemsTitle = POs1.length + " items incomplete"; }
    if (POs1.length > 1) {
        incompleteItemsMessage = POs1.length + " items out of " + (iIndices.length + POs1.length) + " are incomplete."
        + "\nYou can proceed with the " + POorPR + " without these items (OK) or CANCEL and fill in the missing information. The row numbers that need completing are: " + POs1
        + "\n(If you choose CANCEL a list of these row numbers will appear in the bottom right of the window as a reminder)";
    }

    if (POs1.length == 1) { incompleteItemsTitle = POs1.length + " item incomplete"; }
    if (POs1.length == 1) {
        incompleteItemsMessage = POs1.length + " item out of " + (iIndices.length + POs1.length) + " is incomplete."
        + "\nYou can proceed with the " + POorPR + " without this item (OK) or CANCEL and fill in the missing information. The row number that needs completing is: " + POs1
        + "\n(If you choose CANCEL this row number will appear in the bottom right of the window as a reminder)";
    }

    if (POs1.length > 0) {
        var ignoreIncompleteItems = UI.alert(incompleteItemsTitle, incompleteItemsMessage, UI.ButtonSet.OK_CANCEL);
        // Process the user's response.
        if (ignoreIncompleteItems == UI.Button.OK) {
            var toastMsg2 = (iIndices.length == 1) ? "Creating " + POorPR + " with remaining valid item." : "Creating " + POorPR + " with remaining " + iIndices.length + " valid items.";
            SS.toast(toastMsg2, "Step 4 of 5", 60);
            return true;
        }
        if (ignoreIncompleteItems == UI.Button.CANCEL) {
            var toastMsg3 = (POs1.length == 1) ? "Row to check: " + POs1 : "Rows to check: " + POs1;
            SS.toast(toastMsg3, "Incomplete Items", 60);
            return false;
        }
    }
}
/*
 * End Functions originally in POPR 2.gs
 */

/*
 * Functions originally in PR3 and 4.gs

 */

/* * * * * * *
* STAGE 3 of PR Generation Materials - create the PR
*  
* * * * * * */

function PRs3_createPR() {

    var prMain = 'Import';
    var prMain_StartRow = 2;

    var newPRNumber = nextPOPRNumber();
    var newPRName = PROJ_NAME() + ' (' + PROJ_NUMBER() + ') BPR ' + newPRNumber;
    if (isEmergency) { newPRName += ' Emergency Order ' + branchPONumber; }

    var newPRFile = DriveApp.getFileById('1IbNvwMtwtwhje0_quRBEPeHdx7ohDLJ7NhOLZavww64')
        .makeCopy(newPRName,
        DriveApp.getFolderById(PR_FOLDER_ID()));

    newPRFile.setOwner(DRIVE_OWNER());

    var newPRUrl = newPRFile.getUrl();

    var newPRID = newPRUrl.substring(newPRUrl.indexOf('/d/') + 3, newPRUrl.indexOf('/edit'));

    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(newPRID));
    var PR = SpreadsheetApp.getActiveSpreadsheet();
    var PRsh = PR.setActiveSheet(PR.getSheetByName(prMain));

    // cap the number of line items to the PR Items limit
    var li = iIndices.length;
    if (li > PR_ITEMS_LIMIT) { li = PR_ITEMS_LIMIT; }

    // write the retVal value, but add the indices that this PR refers to
    var retVal = { Num: newPRNumber, Url: newPRUrl, ID: newPRID, Name: newPRName, Indices: iIndices.slice(0, li) };

    // get the body area of the PR page
    var mainPRRange = PRsh.getRange(prMain_StartRow, 1, li, 27);
    var mainPRValues = mainPRRange.getValues();

    // get the number of days between today and the delivery date
    var daysToDelivery = dateDiffInDays(poDelivery, new Date());

    var bprPriority = 0;
    if (daysToDelivery < 8) { bprPriority = 1; }

    // fill the main PR page with line items - use shift to start from the top of the list
    // [row][column] but zero indexed not 1 (therefore -1)
    for (var i = 0; i < li; i++) {
        var thisIC = iCode.shift()
    mainPRValues[i][0] = thisIC;
        mainPRValues[i][1] = iDesc.shift();
        mainPRValues[i][2] = iQty.shift();
        mainPRValues[i][3] = iUoM.shift();
        mainPRValues[i][4] = iFactor.shift();
        mainPRValues[i][5] = iBUoM.shift();
        mainPRValues[i][6] = iUnit.shift();
        mainPRValues[i][7] = poDelivery;
        mainPRValues[i][8] = bprPriority;
        mainPRValues[i][9] = ''; // site
        mainPRValues[i][10] = ''; // fill from
        mainPRValues[i][11] = ''; // internal receiving point
        mainPRValues[i][12] = ''; // account number
        mainPRValues[i][13] = ''; // cost centre
        mainPRValues[i][14] = ''; // sub account
        mainPRValues[i][15] = ''; // customer code
        mainPRValues[i][16] = ''; // customer type
        mainPRValues[i][17] = ''; // asset number
        mainPRValues[i][18] = ''; // job number
        mainPRValues[i][19] = PROJ_NUMBER();
        mainPRValues[i][20] = iPDC.shift();
        mainPRValues[i][21] = ''; // supplier code
        mainPRValues[i][22] = ''; // notes to supplier
        mainPRValues[i][23] = iNote.shift();
        mainPRValues[i][24] = ''; // reimbursement customer code
        mainPRValues[i][25] = ''; // reimbursement customer type
        mainPRValues[i][26] = ''; // catalog number
    }

    mainPRRange.setValues(mainPRValues);

    // return to original sheet
    SpreadsheetApp.setActiveSpreadsheet(SS);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SH).setActiveSelection(thisCell);

    // return the details of the new PR to the main script  
    return retVal;

} // enf fn:PRs3_createPR

function PRs4_updateMaterialsList(PRs3) {

    // import the three columns that need updating
    var statusCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_STATUS) + 1), SH.getLastRow(), 1);
    var poNumCol = SH.getRange((MHRI + 2), (wholeList[MHRI].indexOf(H_PONUM) + 1), SH.getLastRow(), 1);
    var poCreatedCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_POCREATED) + 1), SH.getLastRow(), 1);

    var statusValues = statusCol.getValues();
    var poNumFormulas = poNumCol.getFormulas();
    var poCreatedValues = poCreatedCol.getValues();

    var poprLink = '=HYPERLINK(\"' + PRs3.Url + '\",\"' + PRs3.Num + '\")';

    SS.toast("PR " + PRs3.Num + " created, adding links to materials list", "Step 5 of 5", 60);

    // remove the referecnes for the old file for regenerated items
    while (aRegenIndices.length) {
        var r = aRegenIndices.pop();
        statusValues[r][0] = '';
        poNumFormulas[(r - MHRI - 1)][0] = '';
        poCreatedValues[r][0] = '';
    }

    // cap the number of line items to the PR Items limit
    var li = iIndices.length;
    if (li > PR_ITEMS_LIMIT) { li = PR_ITEMS_LIMIT; }

    var _status = null;
    var _statusArray = getCentralDropDowns().ddStatusBPR;
    for (; _statusArray.length > 0;) {
        // get the status dropdown value whose prefix matches the regen prefix
        if (_statusArray[0].substr(0, 1) == REGEN_PREFIX) { _status = _statusArray[0]; }
        _statusArray.shift();
    }

    // enter the values for the newly created file into the correct rows.
    // - use shift to start from the top of the list
    // [row][column] but zero indexed not 1 (therefore -1)
    for (var i = 0; i < li; i++) {
        var j = iIndices.shift();
        statusValues[j][0] = _status;
        poNumFormulas[(j - MHRI - 1)][0] = poprLink;
        poCreatedValues[j][0] = new Date();
    }

    statusCol.setValues(statusValues);
    poNumCol.setFormulas(poNumFormulas);
    poCreatedCol.setValues(poCreatedValues);

    return true;
}

/*
 * End Functions originally in PR3 and 4.gs
 */

/*
 * Functions originally in PR email.gs
 */

function PR_Email(BPRs) {

    Logger.log(poSupplier);

    for (; BPRs.length > 0;) {

        var bpr = BPRs.shift();

        var emailSubject = bpr.Name;

        // the date is set to 8am to remove a problem of GMT dates being shown as the day before when the "now" date is BST
        var _poDelivery = Utilities.formatDate(new Date(poDelivery.setHours(8)), "GMT", "E, dd-MMM-yyyy");

        var emailMessage = "Dear " + poSupplier.name
            + ",\n\nPlease find attached our Purchase Request (number " + bpr.Num + ") for the " + PROJ_NAME + " project."
            + "\nWe are requesting delivery " + delONBY + " " + _poDelivery;
        if (TYPE == COREEXTRA) { emailMessage += "\n\nPlease note that this Purchase Request includes items extra to the core list."; }
        emailMessage += "\n\nKind regards,\netc."

    // get the active PO Google Sheets file ID

    var bprFile = getAsExcel(bpr.ID);//DriveApp.getFileById(bpr.ID);
        bprFile.setName(bpr.Name)

    // send the email as the current user

    MailApp.sendEmail(PRemailTo,
            emailSubject,
            emailMessage,
            {
                attachments: [bprFile],//.getAs(MimeType.MICROSOFT_EXCEL)],
                cc: PRemailCC
            }
            );

    } // end for loop

    return true;

} // end fn:PR_Email

// this uses the advance API Drive service, which was enabled in both the Resources>Advanced Google services... menu and the developer api console
function getAsExcel(spreadsheetId) {
    var file = Drive.Files.get(spreadsheetId);
    var url = file.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });
    return response.getBlob();
}

/*
 * End Functions originally in PR email.gs
 */

/*
 * Functions originally in SVS.gs 
 */

function svsEdit(e) {

    var eRangeA1 = e.range.getA1Notation();
    var eValues = e.range.getValues();

    if ((eRangeA1 == ddSVSProcess && eValues[0][0] == 'GO!')) {
        getPRDataReadyForSVS();
        e.range.setValue('Refresh BPR Items');
        SH.getRange(svsLastUpdated).setValue(new Date());
    }

    if (eRangeA1 == ddSVSMatchConfirm && e.value == 'GO!') {
        confirmMatchedItems();
        e.range.setValue('Confirm Matched Items');
        SH.getRange(svsLastUpdated).setValue(new Date());
    }


    if (e.range.getRow() == svsSortRow) {

        var sortDir = null;

        if (eValues[0][0] == 'A > Z') { sortDir = true; }
        if (eValues[0][0] == 'Z > A') { sortDir = false; }

        if (sortDir != null) {
            SH.getRange((svsSH_HR + 1), //starting row index
                1, // starting column index
                SH.getLastRow(), // number of rows to import
                SH.getLastColumn() // number of columns to import
                ).sort({ column: e.range.getColumn(), ascending: sortDir });
        }// end if sort is valid
        e.range.setValue('Sort?');
    } // end if sort row was edited


} // end fn:svsEdit

function svsTester() {
    Logger.log(svsSH.getLastRow());
}

function getPRDataReadyForSVS() {

    var prData = [];

    var matData = matSH.getRange(1, //starting row index  
        1, // starting column index
        matSH.getLastRow(), // number of rows to import
        matSH.getLastColumn() // number of columns to import
        ).getValues();

    // clear out the 'old' PR area
    var oldPRRange = svsSH.getRange((svsSH_HR),//firstBlankPRRow(),
        1,
        (svsSH.getLastRow() - svsSH_HR),
        5);

    var oldPRValues = oldPRRange.getValues();

    prData.push(oldPRValues[0]);

    if (oldPRValues.length > 0) {
        oldPRRange.clearContent();
    }

    // loop through wholeList to extract the correct items
    for (var m = MHRI + 1; m < matData.length; m++) {
        if (
            // the PR number is a valid and this line is expecting an svs
            matData[m][matData[MHRI].indexOf(H_PONUM)] < 1000
            && matData[m][matData[MHRI].indexOf(H_PONUM)] > 0
            && matData[m][matData[MHRI].indexOf(H_STATUS)].substr(0, 1) == EXPECTINGSVS_PREFIX
            ) {
            var _prData = [];
            _prData.push(matData[m][matData[MHRI].indexOf(H_PONUM)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_LINEID)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_IDESC)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_QTY)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_ACTDEL)]);

            prData.push(_prData);

        } // end if svs expected for valid PR
    } // end for loop through materials list

    if (prData) {

        var thisPRRange = svsSH.getRange(svsSH_HR,
            1,
            prData.length,
            prData[0].length);
        thisPRRange.setValues(prData);
    } // end if there any pr files to process

    return prData;

} // end fn:getPRDataReadyForSVS

function confirmMatchedItems() {

    SS.toast("Reviewing PR / SVS Data matches", "SVS Confirm 1 of 4", 30);

    // get the svs "basket" values
    var currentSVSRange = svsSH.getRange(svsSH_HR,
        1,
        (svsSH.getLastRow() - svsSH_HR),
        svsSH.getLastColumn()
        );

    var currentSVSValues = currentSVSRange.getValues();
    //var currentSVSNotes = currentSVSRange.getNotes();

    var lineIDCol = matSH.getRange(matFirst, col(H_LINEID).n, matLast, 1);
    var lineIDValues = lineIDCol.getValues();

    var orderStatusCol = matSH.getRange(matFirst, col(H_STATUS).n, matLast, 1);
    var orderStatusValues = orderStatusCol.getValues();
    var statusSVSReceived = SVS_RECEIVED();

    var branchPOCol = matSH.getRange(matFirst, col(H_BRANCHPO).n, matLast, 1);
    var branchPOValues = branchPOCol.getValues();
    var branchPOFormulas = branchPOCol.getFormulas();
    var branchPONotes = branchPOCol.getNotes();

    var branchLineCol = matSH.getRange(matFirst, col(H_BRANCHLINE).n, matLast, 1);
    var branchLineValues = branchLineCol.getValues();
    var branchLineNotes = branchLineCol.getNotes();

    var branchSupplierCol = matSH.getRange(matFirst, col(H_BRANCHSUPPLIER).n, matLast, 1);
    var branchSupplierValues = branchSupplierCol.getValues();

    var svsToBeMatched = [];
    //var svsNotesToBeMatched = [];
    svsToBeMatched.push(currentSVSValues[0]); // write the header row back in
    //svsNotesToBeMatched.push(currentSVSNotes[0]);

    SS.toast("Getting SVS PDF files", "SVS Confirm 2 of 4", 30);

    // loop through each SVS item on the right hand side
    for (var s = 1; s < currentSVSValues.length; s++) {

        // set the Match ID (materials line id) to ''
        var svsMatchID = '';

        if (currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_SUPPLIER)] != ''
            || currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_NUM)] != ''
            || currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_LINE)] != ''
            ) {
            svsMatchID = currentSVSValues[s][currentSVSValues[0].indexOf(svsMAT_LINE)];
        }

        // loop throught the materials list...
        for (var t = 0; t < lineIDValues.length; t++) {
            //...store this iteration line id...
            var thisMaterialLineID = lineIDValues[t][0];
            //...and compare it against the SVS Match ID, checking it's not blank
            if (svsMatchID == thisMaterialLineID && svsMatchID != '') {

                // if this is matched, write the svs data to the main sheet
                orderStatusValues[t][0] = statusSVSReceived;

                var wtpoNum = currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_NUM)]; // po number
                var wtsvsUrl = 'https://drive.google.com/open?id=0B6duJkKLBboAZVFsMW45eHg1Qms';
                branchPOFormulas[t][0] = '=HYPERLINK(\"' + wtsvsUrl + '\",\"' + wtpoNum + '\")';

                branchPOValues[t][0] = '=HYPERLINK(\"' + wtsvsUrl + '\",\"' + wtpoNum + '\")';

                //branchPONotes[t][0] = currentSVSNotes[s][currentSVSValues[0].indexOf(svsSVS_WTPO_NUM)]; // buyer details as a note

                branchLineValues[t][0] = currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_LINE)]; // po line/seq number
                //branchLineNotes[t][0] = currentSVSNotes[s][currentSVSValues[0].indexOf(svsSVS_DESC)]; // buyer line item comments as a note

                //branchSupplierValues[t][0] = currentSVSNotes[s][currentSVSValues[0].indexOf(svsSVS_QTY)]; // the WT Supplier is stored as a note on the SVS QTY cell

            } // end if match found
        } // end materials list loop
    } // end svs matcher loop 

    SS.toast("Writing PR / SVS Data matches", "SVS Confirm 3 of 4", 30);

    // write matches back to the main materials list
    orderStatusCol.setValues(orderStatusValues);

    //branchPOCol.setValues(branchPOValues);
    branchPOCol.setFormulas(branchPOFormulas);
    //branchPOCol.setNotes(branchPONotes);

    branchLineCol.setValues(branchLineValues);
    //branchLineCol.setNotes(branchLineNotes);

    branchSupplierCol.setValues(branchSupplierValues);

    // update the SVS matcher display, removing matched items from the left and right hand sides

    SS.toast("Refreshing SVS Matcher", "SVS Confirm 4 of 4", 5);

    currentSVSRange.clearContent().clearNote();

    // get a new range to writte the "to be matched" svs items back
    var refreshedSVSRange = svsSH.getRange(svsSH_HR,
        1,
        svsToBeMatched.length,
        svsToBeMatched[0].length
        );

    refreshedSVSRange.setValues(svsToBeMatched);
    //  refreshedSVSRange.setNotes(svsNotesToBeMatched);

    // write the PR data to the left hand side, and store the result for matching
    var thisPRData = getPRDataReadyForSVS();
    /* 
      // makes sure the new values are written to the matcher before running the match check
      SpreadsheetApp.flush();
  
      // run the match check between left and right
      suggestSVSMatches(thisPRData);
  
      SS.toast("SVS ready to match with materials list","SVS Step 4 of 4",5);
    */
} // confirmMatchedItems

/*
Problem as multiple PDFs from SVS may apply to single PR.
Without reading data content, how to know _which_ PDF SVS is for which lines????
*/

function getPDFSVS() {

    SS.toast("Searching for valid SVS files", "SVS Process 1 of 4", 30);

    var SVSFolder = DriveApp.getFolderById(SVS_FOLDER_ID().toDo);
    var SVSFiles = SVSFolder.getFiles();

    var svsPDFPR = [];
    var svsPDFURL = [];

    // loop through all the files, check their file names and get their IDs
    while (SVSFiles.hasNext()) {
        var svsFile = SVSFiles.next();
        var fileName = svsFile.getName();
        var svsProjNum = fileName.substr(4, 3).toString();
        var svsPRNum = fileName.substr(8, 4).toString();

        if (svsProjNum == PROJ_NUMBER()) {

            if (svsFile.getMimeType() == 'application/pdf') {
                svsPDFURL.push(svsFile.getUrl());
                svsPDFPR.push(svsPRNum);
            } // end if file is a pdf, store the name and URL
        } // end if valid project number
    } // end if files are valid
} // end fn:getPDFSVS

/*
 * End Functions originally in SVS.gs
 */