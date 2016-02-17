/**
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
*/
/*
* Global scope variables
*/
var UI = null;
var CL = null;

var clHeaderRow = 6;
var ddPDN = 'C2';
var ddTrade = 'C3';
var ddSubcategory = 'D3';
var ddType = 'E3';
var ddLocation = 'H3';
var ddPackage = 'C4';
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

/*
* End global scope variables
*/
/*
* Functions originally in Code.gs
*/
function onOpen(e) {
    UI = SpreadsheetApp.getUi();

    var statusUpdateSubMenu = UI.createMenu('Update Order Status').addItem('...to Sent', 'updateStatustoSent');

    UI.createMenu('Purchasing Tools').addItem('Show Purchasing Sidebar', 'purchasingSidebar').addSeparator().addSubMenu(statusUpdateSubMenu).addSeparator().addItem('Clean Up Formats, Formulae and Drop-downs', 'matCleanup').addToUi();
    setupCLPicker();
    setupGoodsReceiver();
    matDropDowns();
}

/*
* Settings for the purchasing side bar
*/
function purchasingSidebar() {
    UI = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutputFromFile('SideBarPlain').setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('Purchasing Side Bar').setWidth(250);
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
    Logger.log("Cell Edited = " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getA1Notation());
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
}

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
        }
    }

    // return to the original spreadsheet
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
    _array = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues(); // an array of rows, each an array of columns wholeList[r-1][c-1]

    // use the JSON.stringify method to convert the 2d array into a string for cache storage
    _string = JSON.stringify(_array);

    // check the JSON string length and break it up into smaller strings
    // !*!*!*!*!*!
    /* ! */ var MSL = 100000;

    for (; _string.length > 0;) {
        if (_string.length >= MSL) {
            _stringCaches.push(_string.substr(0, MSL)); // split out the first MSL (e.g. 100,000) charaters
            _string = _string.substr(MSL); // write the rest of the string back
        }
        if (_string.length < MSL && _string.length > 0) {
            _stringCaches.push(_string); // write the whole remaining string
            _string = ''; // then delete the string
        }
    }

    // store the number of sub-caches in the "parent" cache
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
        _array = JSON.parse(_string); // use the JSON.parse method to convert the cached string back into a 2d array
    }

    return _array;
}

function getCL() {
    var _array = [];
    var _string = '';

    // check the "out of function" placeholder to speed up multiple calls to the cache
    if (CL) {
        return CL;
    }

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

    // store current spreadsheet settings
    var origSS = SpreadsheetApp.getActiveSpreadsheet();
    var origSH = origSS.getActiveSheet();
    var origRange = origSH.getActiveRange();

    _array = putCache(CD_WTSuppliers, CD_WTSuppliers);

    // return to the original spreadsheet
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
        _object = JSON.parse(_string); // use the JSON.parse method to convert the cached string back an object of arrays
        return _object;
    }

    // store current spreadsheet settings
    var origSS = SpreadsheetApp.getActiveSpreadsheet();
    var origSH = origSS.getActiveSheet();
    var origRange = origSH.getActiveRange();

    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(CentralData()));
    var CD = SpreadsheetApp.getActiveSpreadsheet();

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
        DRIVE_OWNER: DRIVE_OWNER, TEMPLATE_IDS: TEMPLATE_IDS,
        ddType: ddType, ddTeams: ddTeams,
        ddUoM: ddUoM, ddPDN: ddPDN, ddPDC: ddPDC,
        ddVATRates: ddVATRates, ddWTSupNames: ddWTSupNames,
        ddStatusWT50: ddStatusWT50, ddStatusHire: ddStatusHire, ddStatusBPR: ddStatusBPR
    };

    // use the JSON.stringify method to convert the 2d array into a string for cache storage
    _string = JSON.stringify(_object);

    // store the array in the cache
    CacheService.getDocumentCache().put(_CDdd, _string);

    // return to the original spreadsheet
    SpreadsheetApp.setActiveSpreadsheet(origSS);
    SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);

    return _object;
}

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

    var cells = [
        ddPDN, CDD.ddPDN,
        ddTeamRequested, CDD.ddTeams
    ];

    for (; cells.length > 0;) {
        var cell = cells.shift();
        var dv = SpreadsheetApp.newDataValidation();
        dv.setAllowInvalid(false);
        var _dd = cells.shift();
        dv.requireValueInList(_dd, true);
        SS.getSheetByName(CORELIST_SHEET).getRange(cell).setDataValidation(dv.build());
    }
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

        if (eValues[0][0] == 'A > Z') {
            sortDir = true;
        }
        if (eValues[0][0] == 'Z > A') {
            sortDir = false;
        }

        if (sortDir != null) {
            SH.getRange((clHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn()).sort({ column: e.range.getColumn(), ascending: sortDir });
        }
        e.range.setValue('Sort?');
    }

    if (e.range.getRow() > clHeaderRow) {
        if (e.range.getColumn() == 7) {
            var _PDN = SH.getRange(ddPDN).getValue();

            // the budget range is the same size as the edited range, but 2 columns to the right
            var budgetPDNRange = e.range.offset(0, 4);
            var budgetPDNValues = budgetPDNRange.getValues();

            for (var i = 0; i < eValues.length; i++) {
                if (eValues[i][0] > 0 && budgetPDNValues[i][0] == '') {
                    budgetPDNValues[i][0] = _PDN;
                }
                if (!eValues[i][0]) {
                    budgetPDNValues[i][0] = '';
                }
            }

            budgetPDNRange.setValues(budgetPDNValues);
        }

        // grab everything, including the basket header
        var currentFLRange = SH.getRange(clHeaderRow, 1, SH.getLastRow(), SH.getLastColumn());

        var basketReturn = basketList(currentFLRange.getValues());

        var basket = basketReturn.basket;
        var basketICs = basketReturn.ics;
        var basketDLs = basketReturn.dls;

        var duplicates = checkForDuplicates(basketICs);

        if (duplicates) {
            SH.getRange(dupWarningCell).setValue('Duplicates\nFound');
            SH.getRange(dupResultsCell).setValue(duplicates.join('\n'));
        }
    }
}

// e is the event object from the edit event in Code.gs
function setCLFilters(e) {
    var cellRef = e.range.getA1Notation();
    var cellValue = e.value;
    if (!cellValue) {
        cellValue = '';
    }

    var fl = [];

    /*  var packageVal = SH.getRange(ddPackage).getValue();
    if (!packageVal){packageVal='';}
    var flPackage = filterCLList(getCL(),packageVal,clPackage);
    */
    var viewAllVal = SH.getRange(ddViewAll).getValue();

    var PDCVal = PD(SH.getRange(ddPDN).getValue());
    if (!PDCVal) {
        PDCVal = '';
    }
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
        }
    }

    // >>>> Trade dropdown was edited
    if (cellRef == ddTrade) {
        clearCells([ddSubcategory, ddType]);
        if (cellValue.length > 0) {
            fl = flTrade;
            setCLDropDown(fl, ddSubcategory, clSubcategory);
        }
    }

    // >>>> Subcategory dropdown was edited
    if (cellRef == ddSubcategory) {
        clearCells([ddType]);
        if (cellValue.length > 0) {
            fl = flSub;
            setCLDropDown(fl, ddType, clType);
        }
        if (cellValue.length == 0) {
            fl = flTrade;
            //      setCLDropDown(fl,ddType,clType);
        }
    }

    // >>>> Type dropdown was edited
    if (cellRef == ddType) {
        if (cellValue.length > 0) {
            fl = flType;
        }
        if (cellValue.length == 0) {
            fl = flSub;
        }
    }

    // >>>> ViewAll dropdown was edited
    if (cellRef == ddViewAll) {
        if (cellValue == 'View All') {
            clearCells([ddPDN, ddTrade, ddSubcategory, ddType]);
            fl = getCL();
        }
        if (cellValue != 'View All') {
            setupCLPicker();
        }
    }

    // filter the CL
    displayFilteredCLList(fl);
    SS.toast("Thanks for waiting", "Done", 1);
}

function displayFilteredCLList(fl) {
    // grab everything, including the basket header
    var currentFLRange = SH.getRange(clHeaderRow, 1, SH.getLastRow(), SH.getLastColumn());

    var basketReturn = basketList(currentFLRange.getValues());

    var basket = basketReturn.basket;
    var basketICs = basketReturn.ics;
    var basketDLs = basketReturn.dls;
    var basketPDNs = basketReturn.pdns;

    var duplicates = checkForDuplicates(basketICs);

    if (duplicates) {
        SH.getRange(dupWarningCell).setValue('Duplicates\nFound');
        SH.getRange(dupResultsCell).setValue(duplicates.join('\n'));
    }

    // clear the basket area, except for the header
    SH.getRange((clHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn()).clearContent();

    // dl is the display list. This is fl with filter columns (trade, sub, type, etc.) removed
    var dl = [];
    var dlDataLinks = [];

    var _PDN = SH.getRange(ddPDN).getValue();

    for (var row = 1; row < fl.length; row++) {
        var _IC = fl[row][fl[0].indexOf(clIC)];
        if (basketICs.indexOf(_IC) < 0 || (basketICs.indexOf(_IC) >= 0 && basketPDNs.indexOf(_PDN) < 0)) {
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
        }
    }

    basket.shift(); // remove the header row

    if (basket.length > 0) {
        // get a range equal in row length to the basket. Write even if there is no basket so that the header is written back in.
        SH.getRange((clHeaderRow + 1), 1, basket.length, basket[0].length).setValues(basket); // set this range with the contents of the filtered list.

        SH.getRange((clHeaderRow + 1), 1, basketDLs.length, 1).setFormulas(basketDLs); // set this range with the formula contents of the filtered list.
    }

    // only export the filtered display list if one exists
    if (dl.length > 0) {
        // get a range equal in row length to the filtered list and start below the existing basket
        SH.getRange((clHeaderRow + 1 + basket.length), 1, dl.length, dl[0].length).setValues(dl); // set this range with the contents of the filtered list.

        // write the data links into the first column
        SH.getRange((clHeaderRow + 1 + basket.length), 1, dlDataLinks.length, 1).setFormulas(dlDataLinks); // set this range with the formula contents of the filtered list.
    }

    // auto resize all the basket columns
    //for (var c=1;c<9;c++){SH.autoResizeColumn(c);}
    SH.getRange(1, 1, 1000, 1).setNumberFormat('@STRING@'); // set the first column to plain text
    //  SH.autoResizeColumn(2);
}

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
        }
    }

    // if duplicates were found, insert a header "row"
    if (duplicatesArray.length > 0) {
        var _dA = ['Row', 'Line', 'Item', 'Qty'];
        duplicatesArray.unshift(_dA.join('\t'));
        return duplicatesArray;
    }

    return null;
}

function filterCLList(list, ddValue, ddOption) {
    var subList = [];

    for (var i = 1; i < list.length; i++) {
        var li = list[i][list[0].indexOf(ddOption)].toString();
        if (li.indexOf(ddValue) > -1) {
            subList.push(list[i]);
        }
    }
    subList.unshift(getCL()[0]); // add in the title row at the beginning
    return subList;
}

function setCLDropDown(list, ddCell_A1, ddType) {
    var aOpt = [];

    for (var i = 1; i < list.length; i++) {
        var opt = list[i][list[0].indexOf(ddType)];
        if (aOpt.indexOf(opt) < 0) {
            aOpt.push(opt); // if this type isn't already in the dropdown option list, add it
        }
    }
    aOpt.sort(); // sort alphabetically

    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.requireValueInList(aOpt, true);
    SS.getSheetByName(CORELIST_SHEET).getRange(ddCell_A1).setDataValidation(dv.build());
}

function basketList(list) {
    var subList = [];
    var icList = [];
    var dlList = [];
    var pdnList = [];

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
        }
    }
    subList.unshift(list[0]); // add in the title row at the beginning
    var retVal = { basket: subList, ics: icList, dls: dlList, pdns: pdnList };
    return retVal;
}

function sendBasket(e) {
    // store the last column index value so that it's not lost after range.clear()
    var clLastCol = SH.getLastColumn();

    // grab everything, including the basket header
    var currentFLRange = SH.getRange(clHeaderRow, 1, SH.getLastRow(), clLastCol);

    var basketReturn = basketList(currentFLRange.getValues());

    var basket = basketReturn.basket;
    var basketICs = basketReturn.ics;
    var basketDLs = basketReturn.dls;

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

    _sh = SS.getSheetByName(MATERIALS_SHEET);
    var curMatRange = _sh.getRange(1, 2, _sh.getLastRow(), _sh.getLastColumn());

    var curMat = curMatRange.getValues();
    var blankRow = 0;

    for (var i = curMat.length - 1; i > MHRI; i--) {
        // check for the first blank row
        if ((curMat[i][curMat[MHRI].indexOf(H_IDESC)] != '') || (curMat[i][curMat[MHRI].indexOf(H_ICODE)] != '') || (curMat[i][curMat[MHRI].indexOf(H_TYPE)] != '')) {
            blankRow = i + 2;
            break;
        }
    }

    // basketExport (be) is the range [r][c] of ordered columns for the materials list
    var be = [];
    var bePDN = [];
    var beSupplier = [];
    var beStatus = [];

    for (var b = 1; b < basket.length; b++) {
        var beItem = [];
        var clType = CORELIST;

        // push in the order needed for the basket columns to match the materials tracking sheet
        beItem.push(basket[b][basket[0].indexOf(bID)]);
        beItem.push(basket[b][basket[0].indexOf(bIC)]);
        if (beItem[1].indexOf('(') > -1) {
            clType = COREEXTRA;
        }
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
        bePDN.push(_bePDN);

        var _beSupplier = [];
        _beSupplier.push(PR_SUPPLIER);
        beSupplier.push(_beSupplier);

        var _beStatus = [];
        _beStatus.push('1 To Be Reviewed');
        beStatus.push(_beStatus);
    }

    _sh.getRange(blankRow, col(H_IDESC).n, be.length, be[0].length).setValues(be);

    _sh.getRange(blankRow, col(H_PDN).n, bePDN.length, bePDN[0].length).setValues(bePDN);

    _sh.getRange(blankRow, col(H_SUPPLIER).n, beSupplier.length, beSupplier[0].length).setValues(beSupplier);

    _sh.getRange(blankRow, col(H_STATUS).n, beStatus.length, beStatus[0].length).setValues(beStatus);

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
        }
    }
    return retVal;
}

function clearCells(cellsA1) {
    for (; cellsA1.length > 0;) {
        SH.getRange(cellsA1.pop()).clearContent().clearDataValidations();
    }
}

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
        MATLIST = matSH.getRange((MHRI + 1), 1, matSH.getLastRow(), matSH.getLastColumn()).getValues();
    }
    return MATLIST;
}

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

        if (eValues[0][0] == 'A > Z') {
            sortDir = true;
        }
        if (eValues[0][0] == 'Z > A') {
            sortDir = false;
        }

        if (sortDir != null) {
            SH.getRange((clHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn()).sort({ column: e.range.getColumn(), ascending: sortDir });
        }
        e.range.setValue('Sort?');
    }

    if (e.range.getRow() > grHeaderRow) {
        if (e.range.getColumn() == 12) {
            var thisDate = new Date();

            // the budget range is the same size as the edited range, but 2 columns to the right
            var dateRange = e.range.offset(0, 1);
            var dateValues = dateRange.getValues();

            for (var i = 0; i < eValues.length; i++) {
                if (eValues[i][0] > 0 && dateValues[i][0] == '') {
                    dateValues[i][0] = thisDate;
                }
                if (!eValues[i][0]) {
                    dateValues[i][0] = '';
                }
            }

            dateRange.setValues(dateValues);
        }
    }
}

// e is the event object from the edit event in Code.gs
function setGRFilters() {
    var fl = [];

    var purchasingRouteVal = SH.getRange(ddPurchasingRoute).getValue();
    if (!purchasingRouteVal) {
        purchasingRouteVal = '';
    }
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

    SH.getRange((clHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn()).sort({ column: 8, ascending: true });

    SS.toast("Thanks for waiting", "Done", 1);
}

function clearDisplayGRList() {
    // grab everything below the header
    var currentFLRange = SH.getRange((grHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn());

    var currentFLValues = currentFLRange.getValues();

    // check if anything has been marked goods received
    currentFLRange.clearContent(); // clear the filter view
}

function compileFilteredGRList(fl, isCore) {
    // dl is the display list. This is fl with filter columns removed
    var dl = [];
    var dlDataLinks = [];

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
            dlItem.push('PR: ' + fl[row][fl[0].indexOf(H_PONUM)] + ' /PO: ' + fl[row][fl[0].indexOf(H_BRANCHPO)]); // col 7

            // get the data link, if available
            var _dataLink = searchCLbyIC(_IC, clDataLink);
            dlDataLink.push('=HYPERLINK(\"' + _dataLink + '\",\"' + _IC + '\")');
            dlDataLinks.push(dlDataLink);
        }

        if (!isCore) {
            dlItem.push(''); // col 4, no manufacturer for WT50
            dlItem.push(''); // col 5, no part number for WT50
            dlItem.push(fl[row][fl[0].indexOf(H_SUPPLIER)]); // col 6
            dlItem.push(WT_PREFIX + ' ' + PROJ_NUMBER() + ' ' + fl[row][fl[0].indexOf(H_PONUM)]); // col 7
        }

        dlItem.push(fl[row][fl[0].indexOf(H_ACTDEL)]); // col 8
        dlItem.push(fl[row][fl[0].indexOf(H_QTY)]); // col 9
        dlItem.push(fl[row][fl[0].indexOf(H_QTYLEFT)]); // col 10
        dlItem.push(fl[row][fl[0].indexOf(H_PUOM)]); // col 11

        dl.push(dlItem);
    }

    if (dl.length > 0) {
        // get a range equal in row length to the filtered list and write anyway to reinsert the header
        SH.getRange((grHeaderRow + 1), 1, dl.length, dl[0].length).setValues(dl); // set this range with the contents of the filtered list.

        if (isCore) {
            // write the data links into the first column, if the purchasing route isCore
            SH.getRange((grHeaderRow + 1), 2, dlDataLinks.length, 1).setFormulas(dlDataLinks); // set this range with the formula contents of the filtered list.
        }
    }
    // auto resize selected columns
    /*  SH.autoResizeColumn(3);
    if (isCore){
    SH.autoResizeColumn(4);
    SH.autoResizeColumn(5);
    }
    SH.autoResizeColumn(6);
    */
}

function filterGRList(list, ddValue, ddOption) {
    var subList = [];

    for (var i = (1); i < list.length; i++) {
        // check that the status is "pending delivery" and items are left to be delivered
        if (list[i][list[0].indexOf(H_STATUS)].toString().substr(0, 1) > 3 && list[i][list[0].indexOf(H_STATUS)].toString().substr(0, 1) < 7 && list[i][list[0].indexOf(H_QTYLEFT)] > 0) {
            var li = list[i][list[0].indexOf(ddOption)];
            if (li.indexOf(ddValue) > -1) {
                subList.push(list[i]);
            }
        }
    }
    subList.unshift(getML()[0]); // add in the title row at the beginning
    return subList;
}

function setGRDropDown(list, ddCell_A1, ddType) {
    var aOpt = [];

    for (var i = 1; i < list.length; i++) {
        var opt = list[i][list[0].indexOf(ddType)];
        if (aOpt.indexOf(opt) < 0) {
            aOpt.push(opt); // if this type isn't already in the dropdown option list, add it
        }
    }
    aOpt.sort(); // sort alphabetically

    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.requireValueInList(aOpt, true);
    SS.getSheetByName(GOODS_RECEIVING_SHEET).getRange(ddCell_A1).setDataValidation(dv.build());
}

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
        if (grQTYRcvdVals[g][0] > 0) {
            // get the index number from the materials sheet for this lineID
            var mtIndex = mtLineIDVals.indexOf(grLineIDVals[g][0]);
            Logger.log(mtIndex);
            mtQTYRcvdVals[mtIndex][0] = grQTYRcvdVals[g][0];
            mtDateRcvdVals[mtIndex][0] = grDateRcvdVals[g][0];
            mtNotesVals[mtIndex][0] = grNotesVals[g][0];

            if (grQTYRcvdVals[g][0] < grQTYRemVals[g][0]) {
                mtStatusVals[mtIndex][0] = '6 Part Delivered';
            }
            if (grQTYRcvdVals[g][0] >= grQTYRemVals[g][0]) {
                mtStatusVals[mtIndex][0] = '7 Delivered';
            }
        }
    }

    SS.toast("Writing Delivery Data", "Goods Receiving 3 of 3", 600);

    mtQTYRcvdRange.setValues(mtQTYRcvdVals);
    mtDateRcvdRange.setValues(mtDateRcvdVals);
    mtNotesRange.setValues(mtNotesVals);
    mtStatusRange.setValues(mtStatusVals);

    SS.toast("Done", "Goods Received", 5);

    grSH.getRange(ddMarkReceived).setValue('Mark Goods Received');

    grSH.getRange(ddPurchasingRoute).setValue('GO!');
    setGRFilters();
}
/*
* End Functions originally in GoodsReceiving.gs
*/
//# sourceMappingURL=CentralCode.js.map
