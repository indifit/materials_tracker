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
var clSortRow = 5;
var clSubcategory = 'Product sub category';
var clType = 'Type';
var clLocation = 'Location';
var clPackage = 'Package key';
var clPDC = 'PDC';
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
var poSupplierName = '';
var poDelivery = '';
var poNumber = '';
var requestingDept = '';
var sameDept = '';
var delONBY = '';
var wholeRange = '';
var wholeList = [];
var aCD_WTSuppliers = [];
var iQty = [];
var iUoM = [];
var iFactor = [];
var iBUoM = [];
var iCode = [];
var iDesc = [];
var iUnit = [];
var iPDN = [];
var iPDC = [];
var iIndices = [];
var iNote = [];
var aRegenIndices = [];
var isEmergency = false;
var branchPONumber = '';
var svsSH = SS.getSheetByName(SVS_MATCHER_SHEET);
var svsSH_HR = 6;
var svsSortRow = 5;
var ddSVSProcess = 'B2';
var ddSVSMatchConfirm = 'H2';
var svsLastUpdated = 'C3';
var svsMAT_PR = 'PR Number';
var svsMAT_LINE = 'Line ID';
var svsMAT_DESC = 'Item Description';
var svsMAT_QTY = 'Item Quantity';
var svsMAT_DEL = 'Requested Delivery Date';
var svsSVS_SUPPLIER = 'SVS Supplier';
var svsSVS_WTPO_NUM = 'SVS WT PO Number';
var svsSVS_WTPO_LINE = 'SVS WT PO Line';
function onOpen(e) {
    UI = SpreadsheetApp.getUi();
    var statusUpdateSubMenu = UI.createMenu('Update Order Status')
        .addItem('...to Sent', 'updateStatustoSent');
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
function purchasingSidebar() {
    UI = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutputFromFile('SideBarPlain')
        .setSandboxMode(HtmlService.SandboxMode)
        .setTitle('Purchasing Side Bar')
        .setWidth(250);
    UI.showSidebar(html);
}
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
}
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
    var _object;
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
    _string = JSON.stringify(_object);
    CacheService.getDocumentCache().put(_CDdd, _string);
    SpreadsheetApp.setActiveSpreadsheet(origSS);
    SpreadsheetApp.setActiveSheet(origSH).setActiveRange(origRange);
    return _object;
}
function setupCLPicker() {
    getCL();
    var CDD = getCentralDropDowns();
    CDD.ddPDN.pop();
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
function setCLFilters(e) {
    var cellRef = e.range.getA1Notation();
    var cellValue = e.value;
    if (!cellValue) {
        cellValue = '';
    }
    var fl = [];
    var viewAllVal = SH.getRange(ddViewAll).getValue();
    var PDCVal = PD(SH.getRange(ddPDN).getValue());
    if (!PDCVal) {
        PDCVal = '';
    }
    var flPDC = filterCLList(getCL(), PDCVal, clPDC);
    var tradeVal = SH.getRange(ddTrade).getValue();
    var flTrade = filterCLList(flPDC, tradeVal, clTrade);
    var subVal = SH.getRange(ddSubcategory).getValue();
    var flSub = filterCLList(flTrade, subVal, clSubcategory);
    var typeVal = SH.getRange(ddType).getValue();
    var flType = filterCLList(flSub, typeVal, clType);
    if (cellRef == ddPDN) {
        clearCells([ddTrade, ddSubcategory, ddType]);
        if (cellValue.length > 0) {
            fl = flPDC;
            setCLDropDown(fl, ddTrade, clTrade);
        }
    }
    if (cellRef == ddTrade) {
        clearCells([ddSubcategory, ddType]);
        if (cellValue.length > 0) {
            fl = flTrade;
            setCLDropDown(fl, ddSubcategory, clSubcategory);
        }
    }
    if (cellRef == ddSubcategory) {
        clearCells([ddType]);
        if (cellValue.length > 0) {
            fl = flSub;
            setCLDropDown(fl, ddType, clType);
        }
        if (cellValue.length == 0) {
            fl = flTrade;
        }
    }
    if (cellRef == ddType) {
        if (cellValue.length > 0) {
            fl = flType;
        }
        if (cellValue.length == 0) {
            fl = flSub;
        }
    }
    if (cellRef == ddViewAll) {
        if (cellValue == 'View All') {
            clearCells([ddPDN, ddTrade, ddSubcategory, ddType]);
            fl = getCL();
        }
        if (cellValue != 'View All') {
            setupCLPicker();
        }
    }
    displayFilteredCLList(fl);
    SS.toast("Thanks for waiting", "Done", 1);
}
function displayFilteredCLList(fl) {
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
    SH.getRange((clHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn()).clearContent();
    var dl = [];
    var dlDataLinks = [];
    var _PDN = SH.getRange(ddPDN).getValue();
    for (var row = 1; row < fl.length; row++) {
        var _IC = fl[row][fl[0].indexOf(clIC)];
        if (basketICs.indexOf(_IC) < 0
            || (basketICs.indexOf(_IC) >= 0 && basketPDNs.indexOf(_PDN) < 0)) {
            var dlItem = [];
            var dlDataLink = [];
            dlItem.push(_IC);
            dlItem.push(fl[row][fl[0].indexOf(clID)]);
            dlItem.push(fl[row][fl[0].indexOf(clMfg)]);
            dlItem.push(fl[row][fl[0].indexOf(clBrand)]);
            dlItem.push(fl[row][fl[0].indexOf(clPartNo)]);
            dlItem.push(fl[row][fl[0].indexOf(clLastPrice)]);
            dlItem.push('');
            dlItem.push(fl[row][fl[0].indexOf(clPUoM)]);
            dlItem.push(fl[row][fl[0].indexOf(clFactor)]);
            dlItem.push(fl[row][fl[0].indexOf(clBUoM)]);
            dl.push(dlItem);
            var _dataLink = fl[row][fl[0].indexOf(clDataLink)];
            dlDataLink.push('=HYPERLINK(\"' + _dataLink + '\",\"' + _IC + '\")');
            dlDataLinks.push(dlDataLink);
        }
    }
    basket.shift();
    if (basket.length > 0) {
        SH.getRange((clHeaderRow + 1), 1, basket.length, basket[0].length).setValues(basket);
        SH.getRange((clHeaderRow + 1), 1, basketDLs.length, 1).setFormulas(basketDLs);
    }
    if (dl.length > 0) {
        SH.getRange((clHeaderRow + 1 + basket.length), 1, dl.length, dl[0].length).setValues(dl);
        SH.getRange((clHeaderRow + 1 + basket.length), 1, dlDataLinks.length, 1).setFormulas(dlDataLinks);
    }
    SH.getRange(1, 1, 1000, 1).setNumberFormat('@STRING@');
}
function checkForDuplicates(basketICs) {
    SH.getRange(dupWarningCell + ':' + dupResultsCell).clearContent();
    var tempMatSh = SS.getSheetByName(MATERIALS_SHEET);
    var tempMatList = tempMatSh.getRange(1, 1, tempMatSh.getLastRow(), 9).getValues();
    var duplicatesArray = [];
    for (var t = (tempMatList.length - 1); t > 0; t--) {
        var thisIC = tempMatList[t][tempMatList[MHRI].indexOf(H_ICODE)];
        var thisICasString = thisIC.toString();
        if (basketICs.indexOf(thisIC) > -1 || basketICs.indexOf(thisICasString) > -1) {
            var dA = [
                (t + 1),
                (tempMatList[t][tempMatList[MHRI].indexOf(H_LINEID)]),
                thisIC,
                (tempMatList[t][tempMatList[MHRI].indexOf(H_QTY)])
            ];
            duplicatesArray.unshift(dA.join('\t'));
        }
    }
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
    subList.unshift(getCL()[0]);
    return subList;
}
function setCLDropDown(list, ddCell_A1, ddType) {
    var aOpt = [];
    for (var i = 1; i < list.length; i++) {
        var opt = list[i][list[0].indexOf(ddType)];
        if (aOpt.indexOf(opt) < 0) {
            aOpt.push(opt);
        }
    }
    aOpt.sort();
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
            subList.push(list[i]);
            pdnList.push(list[i][list[0].indexOf(bPDN)]);
            var _ic = list[i][list[0].indexOf(bIC)].toString();
            icList.push(_ic);
            var _dataLink = searchCLbyIC(_ic, clDataLink);
            var _dlFormula = [];
            _dlFormula.push('=HYPERLINK(\"' + _dataLink + '\",\"' + _ic + '\")');
            dlList.push(_dlFormula);
        }
    }
    subList.unshift(list[0]);
    var retVal = { basket: subList, ics: icList, dls: dlList, pdns: pdnList };
    return retVal;
}
function sendBasket(e) {
    var clLastCol = SH.getLastColumn();
    var currentFLRange = SH.getRange(clHeaderRow, 1, SH.getLastRow(), clLastCol);
    var basketReturn = basketList(currentFLRange.getValues());
    var basket = basketReturn.basket;
    var basketICs = basketReturn.ics;
    var basketDLs = basketReturn.dls;
    var duplicates = checkForDuplicates(basketICs);
    Logger.log('tested for duplicates');
    Logger.log('basket size is: ' + (basket.length - 1));
    if (basket.length == 1 && e.value == 'GO!') {
        e.range.setValue("Add to Materials Tracker");
        SS.toast(".", ".", 1);
        UI.alert("No Items to Add", "There were no basket items found to add to the materials tracker", UI.ButtonSet.OK);
        return;
    }
    currentFLRange.clearContent();
    var basketHeader = SH.getRange(clHeaderRow, 1, 1, clLastCol);
    basketHeader.setValues([basket[0]]);
    _sh = SS.getSheetByName(MATERIALS_SHEET);
    var curMatRange = _sh.getRange(1, 2, _sh.getLastRow(), _sh.getLastColumn());
    var curMat = curMatRange.getValues();
    var blankRow = 0;
    for (var i = curMat.length - 1; i > MHRI; i--) {
        if ((curMat[i][curMat[MHRI].indexOf(H_IDESC)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_ICODE)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_TYPE)] != '')) {
            blankRow = i + 2;
            break;
        }
    }
    var be = [];
    var bePDN = [];
    var beSupplier = [];
    var beStatus = [];
    for (var b = 1; b < basket.length; b++) {
        var beItem = [];
        var clType = CORELIST;
        beItem.push(basket[b][basket[0].indexOf(bID)]);
        beItem.push(basket[b][basket[0].indexOf(bIC)]);
        if (beItem[1].indexOf('(') > -1) {
            clType = COREEXTRA;
        }
        beItem.push(clType);
        beItem.push('');
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
    e.range.setValue("Add to Materials Tracker");
    clearCells([ddPDN, ddTrade, ddSubcategory, ddType]);
    setupCLPicker();
    SS.toast("Thanks for waiting", "Done", 1);
}
function searchCLbyIC(IC, searchHeader) {
    CL = getCL();
    var retVal = null;
    for (var i = 0; i < CL.length; i++) {
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
function setupGoodsReceiver() {
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
        flWT.shift();
        fl = flCore.concat(flWT);
    }
    Logger.log(flPurchasingRoute);
    clearDisplayGRList();
    var isCore = true;
    compileFilteredGRList(fl, true);
    SH.getRange(grLastUpdated).setValue(new Date());
    SH.getRange(ddPurchasingRoute).setValue('Refresh Items');
    SH.getRange((clHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn()).sort({ column: 8, ascending: true });
    SS.toast("Thanks for waiting", "Done", 1);
}
function clearDisplayGRList() {
    var currentFLRange = SH.getRange((grHeaderRow + 1), 1, SH.getLastRow(), SH.getLastColumn());
    var currentFLValues = currentFLRange.getValues();
    currentFLRange.clearContent();
}
function compileFilteredGRList(fl, isCore) {
    var dl = [];
    var dlDataLinks = [];
    for (var row = 1; row < fl.length; row++) {
        var _IC = fl[row][fl[0].indexOf(H_ICODE)];
        var dlItem = [];
        var dlDataLink = [];
        dlItem.push(fl[row][fl[0].indexOf(H_LINEID)]);
        dlItem.push(_IC);
        dlItem.push(fl[row][fl[0].indexOf(H_IDESC)]);
        if (isCore) {
            dlItem.push(searchCLbyIC(_IC, clMfg));
            dlItem.push(searchCLbyIC(_IC, clPartNo));
            dlItem.push(fl[row][fl[0].indexOf(H_BRANCHSUPPLIER)]);
            dlItem.push('PR: ' + fl[row][fl[0].indexOf(H_PONUM)]
                + ' /PO: ' + fl[row][fl[0].indexOf(H_BRANCHPO)]);
            var _dataLink = searchCLbyIC(_IC, clDataLink);
            dlDataLink.push('=HYPERLINK(\"' + _dataLink + '\",\"' + _IC + '\")');
            dlDataLinks.push(dlDataLink);
        }
        if (!isCore) {
            dlItem.push('');
            dlItem.push('');
            dlItem.push(fl[row][fl[0].indexOf(H_SUPPLIER)]);
            dlItem.push(WT_PREFIX + ' ' + PROJ_NUMBER() + ' ' + fl[row][fl[0].indexOf(H_PONUM)]);
        }
        dlItem.push(fl[row][fl[0].indexOf(H_ACTDEL)]);
        dlItem.push(fl[row][fl[0].indexOf(H_QTY)]);
        dlItem.push(fl[row][fl[0].indexOf(H_QTYLEFT)]);
        dlItem.push(fl[row][fl[0].indexOf(H_PUOM)]);
        dl.push(dlItem);
    }
    if (dl.length > 0) {
        SH.getRange((grHeaderRow + 1), 1, dl.length, dl[0].length).setValues(dl);
        if (isCore) {
            SH.getRange((grHeaderRow + 1), 2, dlDataLinks.length, 1).setFormulas(dlDataLinks);
        }
    }
}
function filterGRList(list, ddValue, ddOption) {
    var subList = [];
    for (var i = (1); i < list.length; i++) {
        if (list[i][list[0].indexOf(H_STATUS)].toString().substr(0, 1) > 3
            && list[i][list[0].indexOf(H_STATUS)].toString().substr(0, 1) < 7
            && list[i][list[0].indexOf(H_QTYLEFT)] > 0) {
            var li = list[i][list[0].indexOf(ddOption)];
            if (li.indexOf(ddValue) > -1) {
                subList.push(list[i]);
            }
        }
    }
    subList.unshift(getML()[0]);
    return subList;
}
function setGRDropDown(list, ddCell_A1, ddType) {
    var aOpt = [];
    for (var i = 1; i < list.length; i++) {
        var opt = list[i][list[0].indexOf(ddType)];
        if (aOpt.indexOf(opt) < 0) {
            aOpt.push(opt);
        }
    }
    aOpt.sort();
    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.requireValueInList(aOpt, true);
    SS.getSheetByName(GOODS_RECEIVING_SHEET).getRange(ddCell_A1).setDataValidation(dv.build());
}
function markGoodsReceived() {
    SS.toast("Collecting Materials and Delivery Data", "Goods Receiving 1 of 3", 600);
    var grSH = SS.getSheetByName(GOODS_RECEIVING_SHEET);
    var grSHlr = grSH.getLastRow();
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
    var mtLineIDRange = matSH.getRange(matFirst, col(H_LINEID).n, matLast, 1);
    var mtLineIDVals = mtLineIDRange.getValues();
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
var matSH = SS.getSheetByName(MATERIALS_SHEET);
var matFirst = MHRI + 2;
var matLast = matSH.getLastRow();
function matEdit(e) {
    if (e.range.getColumn() == col(H_SUPPLIER).n) {
        e.range.clearNote();
        matSH.getRange(col(H_ADMINCODE).a + e.range.getRow()).setValue('');
        var supplier = getSupplier(e.value);
        if (supplier.name == e.value) {
            var hoverNote = supplier.contact + '\n' + supplier.tel + '\n' + supplier.email;
            e.range.setNote(hoverNote);
            matSH.getRange(col(H_ADMINCODE).a + e.range.getRow()).setValue(supplier.admin);
        }
    }
    if (e.range.getColumn() == col(H_STATUS).n) {
        var thisPDN = matSH.getRange(col(H_PDN).a + e.range.getRow());
        var thisPDC = matSH.getRange(col(H_PDC).a + e.range.getRow());
        if (e.value.substr(0, 1) >= VOID_PREFIX) {
            if (thisPDN.getNote() == '') {
                thisPDN.setNote('was: ' + thisPDN.getValue());
                thisPDC.setNote('was: ' + thisPDC.getValue());
                thisPDN.setValue('void');
            }
        }
        if (e.value.substr(0, 1) < VOID_PREFIX) {
            if (thisPDN.getNote() != '') {
                thisPDN.setValue(thisPDN.getNote().substr(5));
                thisPDN.clearNote();
                thisPDC.clearNote();
            }
        }
    }
}
function matCleanup() {
    matDropDowns();
    matDateFormat();
    matNumberFormat();
    matSetFormulae();
}
function matDropDowns() {
    var CDD = getCentralDropDowns();
    var allSuppliers = getSupplier('*all*');
    var supplierNames = [];
    for (; allSuppliers.length > 0;) {
        supplierNames.unshift(allSuppliers.pop()[0]);
    }
    supplierNames.shift();
    var pdnDDSource = CDD.ddPDN;
    var columns = [
        H_TYPE, CDD.ddType,
        H_TEAM, CDD.ddTeams,
        H_PUOM, CDD.ddUoM,
        H_BUOM, CDD.ddUoM,
        H_PDN, pdnDDSource,
        H_VATRATE, CDD.ddVATRates,
        H_SUPPLIER, supplierNames,
        H_EMERGENCY, ['Yes', 'No'],
        H_STATUS, CDD.ddStatusBPR,
        H_HIRESTATUS, CDD.ddStatusHire
    ];
    for (; columns.length > 0;) {
        var thisCol = columns.shift();
        var colRange = col(thisCol).a + (MHRI + 2) + ':' + col(thisCol).a;
        var dv = SpreadsheetApp.newDataValidation();
        dv.setAllowInvalid(false);
        var _dd = columns.shift();
        dv.requireValueInList(_dd, true);
        SS.getSheetByName(MATERIALS_SHEET).getRange(colRange).setDataValidation(dv.build());
    }
}
function matDateFormat() {
    var dtColumns = [
        col(H_ACTDEL).a,
        col(H_OFF).a,
        col(H_POCREATED).a
    ];
    for (; dtColumns.length > 0;) {
        var _col = dtColumns.pop();
        var fRange = matSH.getRange(_col + matFirst + ":" + _col + matLast);
        var fDT = [];
        var fD = [];
        for (var i = fRange.getNumRows(); i > 0; i--) {
            fDT.push(["dd/MM/yyyy HH:mm:ss"]);
            fD.push(["dd/MM/yyyy"]);
        }
        fRange.setNumberFormats(fDT);
        fRange.setNumberFormats(fD);
        fRange.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
    }
}
function matNumberFormat() {
    var colNumFormat = [
        col(H_ICODE).a, '@STRING@',
        col(H_UNIT).a, '0.00',
        col(H_NET).a, '0.00',
        col(H_VATRATE).a, '0%',
        col(H_VATVALUE).a, '0.00',
        col(H_LINEVALUE).a, '0.00'
    ];
    for (; colNumFormat.length > 0;) {
        var _col = colNumFormat.shift();
        var _format = colNumFormat.shift();
        var fRange = matSH.getRange(_col + matFirst + ":" + _col + matLast);
        var fNum = [];
        for (var i = fRange.getNumRows(); i > 0; i--) {
            fNum.push([_format]);
        }
        fRange.setNumberFormats(fNum);
    }
}
function matSetFormulae() {
    var formulaColumns = [
        col(H_NET).a,
        '=IF(AND(EQ(' + col(H_QTY).a + matFirst + ',""),EQ(' + col(H_UNIT).a + matFirst + ',"")),"",' + col(H_QTY).a + matFirst + '*' + col(H_UNIT).a + matFirst + ')',
        col(H_VATVALUE).a,
        '=IF(EQ(' + col(H_VATRATE).a + matFirst + ',""),"",' + col(H_NET).a + matFirst + '*(RIGHT(' + col(H_VATRATE).a + matFirst + ',(LEN(' + col(H_VATRATE).a + matFirst + ')-2))))',
        col(H_LINEVALUE).a,
        '=IF(AND(EQ(' + col(H_NET).a + matFirst + ',"")),"",' + col(H_NET).a + matFirst + '+' + col(H_VATVALUE).a + matFirst + ')',
        col(H_PDC).a,
        '=IF(EQ(' + col(H_PDN).a + matFirst + ',""),"",VLOOKUP(' + col(H_PDN).a + matFirst + ',PD_Range,2,FALSE))',
        col(H_QTYLEFT).a,
        '=IF(EQ(' + col(H_QTY).a + matFirst + ',""),"",' + col(H_QTY).a + matFirst + '-' + col(H_QTYRCVD).a + matFirst + ')'
    ];
    for (; formulaColumns.length > 0;) {
        var initCol = formulaColumns.shift();
        var initRange = matSH.getRange(initCol + matFirst);
        initRange.setFormula(formulaColumns.shift());
        initRange.copyTo(matSH.getRange(initCol + (matFirst + 1) + ":" + initCol + matLast));
    }
}
function col(columnHeader) {
    if (!hRow) {
        hRow = matSH.getRange((MHRI + 1), 1, 1, matSH.getLastColumn()).getValues();
    }
    var hIndex = hRow[0].indexOf(columnHeader);
    if (hIndex > -1) {
        var retVal = { i: hIndex, n: (hIndex + 1), a: alphaCols[hIndex] };
        return retVal;
    }
}
function updateStatustoSent() {
    UI = SpreadsheetApp.getUi();
    var thisRowIndex = SH.getActiveCell().getRow() - matFirst;
    if (thisRowIndex <= MHRI) {
        var wrongRow = UI.alert("Status Update Error", "You have currently selected one of the header rows."
            + "\nPlease select a cell in a row from the materials list.", UI.ButtonSet.OK);
        return false;
    }
    if (SH.getName() != MATERIALS_SHEET) {
        var uiResponse = UI.alert("Status Update Error", "You are not on the " + MATERIALS_SHEET + " sheet.", UI.ButtonSet.OK);
        return false;
    }
    var statusBPRReady = '3 BPR Prepared';
    var statusBPRSent = '4 BPR Sent';
    var mtPRNoRange = matSH.getRange(matFirst, col(H_PONUM).n, matLast, 1);
    var mtPRNoVals = mtPRNoRange.getValues();
    var mtStatusRange = matSH.getRange(matFirst, col(H_STATUS).n, matLast, 1);
    var mtStatusVals = mtStatusRange.getValues();
    var thisBPRNo = mtPRNoVals[thisRowIndex][0];
    var okToUpdate = true;
    for (var u = 0; u < mtPRNoVals.length; u++) {
        if (mtPRNoVals[u][0] == thisBPRNo) {
            if (mtStatusVals[u][0] != statusBPRReady) {
                okToUpdate = false;
            }
            if (mtStatusVals[u][0] == statusBPRReady) {
                mtStatusVals[u][0] = statusBPRSent;
            }
        }
        if (!okToUpdate) {
            break;
        }
    }
    if (!okToUpdate) {
        var uiResponse = UI.alert("Status Update Error", "Not all line items for BPR " + thisBPRNo + " have a status of \"" + statusBPRReady + "\".", UI.ButtonSet.OK);
        return false;
    }
    if (okToUpdate) {
        mtStatusRange.setValues(mtStatusVals);
    }
}
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
    for (var n = 0; n < ncValues.length; n++) {
        var _ncToAdd = [];
        var _ncIC = '';
        var _ncID = '';
        var _ncType = COREEXTRA;
        var _ncTemp = '';
        if (ncValues[n][0] != '') {
            _ncTemp = ncValues[n][0].toString().trim();
            _ncIC = '(' + _ncTemp + ')';
            _ncID = _ncIC;
        }
        if (ncValues[n][1] != '') {
            _ncTemp = ncValues[n][1].toString().trim().toUpperCase();
            if (_ncID != '') {
                _ncID += ' ';
            }
            _ncID += _ncTemp;
        }
        if (ncValues[n][2] != '') {
            _ncTemp = ncValues[n][2].toString().trim().toUpperCase();
            if (_ncID != '') {
                _ncID += ' ';
            }
            _ncID += _ncTemp;
        }
        if (ncValues[n][3] != '') {
            _ncTemp = ncValues[n][3].toString().trim().toUpperCase();
            if (_ncID != '') {
                _ncID += ' ';
            }
            _ncID += _ncTemp;
        }
        if (ncValues[n][4] != '') {
            _ncTemp = ncValues[n][4].toString().trim().toUpperCase();
            if (_ncID != '') {
                _ncID += ' ';
            }
            _ncID += _ncTemp + ':';
        }
        if (ncValues[n][5] != '') {
            _ncTemp = ncValues[n][5].toString().trim().toLowerCase();
            if (_ncID != '') {
                _ncID += ' ';
            }
            _ncID += _ncTemp;
        }
        if (_ncID != '') {
            _ncToAdd.push(_ncID);
            _ncToAdd.push('');
            _ncToAdd.push(_ncType);
            ncToAdd.push(_ncToAdd);
        }
    }
    SS.toast("Adding Non Core Items", "Non Core 2 of 3", 600);
    matSH = SS.getSheetByName(MATERIALS_SHEET);
    var curMatRange = matSH.getRange(1, 2, matSH.getLastRow(), matSH.getLastColumn());
    var curMat = curMatRange.getValues();
    var blankRow = 0;
    for (var i = curMat.length - 1; i > MHRI; i--) {
        if ((curMat[i][curMat[MHRI].indexOf(H_IDESC)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_ICODE)] != '')
            || (curMat[i][curMat[MHRI].indexOf(H_TYPE)] != '')) {
            blankRow = i + 2;
            break;
        }
    }
    matSH.getRange(blankRow, col(H_IDESC).n, ncToAdd.length, ncToAdd[0].length).setValues(ncToAdd);
    SS.toast("Clearing Non Core (Free Text)", "Non Core 3 of 3", 5);
    ncRange.clear();
}
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
        .makeCopy(newPOName + ' ~ ' + poSupplierName, DriveApp.getFolderById(PO_FOLDER_ID));
    newPOFile.setOwner(DRIVE_OWNER);
    var newPOUrl = newPOFile.getUrl();
    var newPOID = newPOUrl.substring(newPOUrl.indexOf('/d/') + 3, newPOUrl.indexOf('/edit'));
    var retVal = { Num: newPONumber, Url: newPOUrl, ID: newPOID, Name: newPOName };
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(newPOID));
    var PO = SpreadsheetApp.getActiveSpreadsheet();
    var POsh = SpreadsheetApp.getActiveSheet();
    var i = iIndices.length;
    var appSubs = [];
    if (i > 10 && i <= 250) {
        var a = Math.floor(((i - 10) / 25) + 1);
        var appNames = [];
        for (var j = 0; j < a; j++) {
            POsh = PO.setActiveSheet(PO.getSheetByName(poAppendix)).copyTo(PO).setName(poAppendix + ' ' + alphaCols[j]);
            POsh.getRange(appPONumberCell).setValue(newPOName);
            POsh.getRange(appNameCell).setValue(POsh.getSheetName());
            appNames.push(POsh.getSheetName());
        }
        var appRow = ((i - 10 + a) % 25);
        for (; appNames.length > 0;) {
            POsh = PO.getSheetByName(appNames.pop()).activate();
            var appendixRange = POsh.getRange(poAppendix_StartRow - 1, 1, appRow + 1, 8);
            var appendixValues = appendixRange.getValues();
            var appendixPDNRange = POsh.getRange(poAppendix_StartRow - 1, 13, appRow + 1);
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
    }
    POsh = PO.setActiveSheet(PO.getSheetByName(poMain));
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
    PO.getRangeByName('ship_adr1').setValue(poShippingInfo[0][0]);
    PO.getRangeByName('ship_adr2').setValue(poShippingInfo[1][0]);
    PO.getRangeByName('ship_adr3').setValue(poShippingInfo[2][0]);
    PO.getRangeByName('ship_adr4').setValue(poShippingInfo[3][0]);
    PO.getRangeByName('ship_adr5').setValue(poShippingInfo[4][0]);
    var mainPORange = POsh.getRange(poMain_StartRow - 1, 1, poMain_EndRow - poMain_StartRow + 2, 8);
    var mainPOValues = mainPORange.getValues();
    var mainPOPDNRange = POsh.getRange(poMain_StartRow - 1, 13, poMain_EndRow - poMain_StartRow + 2);
    var mainPOPDNValues = mainPOPDNRange.getValues();
    for (; iQty.length > 0;) {
        mainPOValues[iQty.length][0] = (10 * (iQty.length));
        mainPOValues[iQty.length][1] = iQty.pop();
        mainPOValues[iUoM.length][3] = iUoM.pop();
        mainPOValues[iDesc.length][4] = iDesc.pop();
        mainPOValues[iUnit.length][7] = iUnit.pop();
        mainPOPDNValues[iPDN.length][0] = iPDN.pop();
    }
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
    PO.deleteSheet(PO.getSheetByName(poAppendix));
    SpreadsheetApp.setActiveSpreadsheet(SS);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SH).setActiveSelection(thisCell);
    return retVal;
}
function nextPOPRNumber() {
    if (poNumber != '') {
        return poNumber;
    }
    var nextNumber = '';
    var upperLimit = 2000;
    var lowerLimit = 0;
    var currentNumbers = [];
    var POPRnumbers = SH.getRange((MHRI + 2), (wholeList[MHRI].indexOf(H_PONUM) + 1), SH.getLastRow(), 1).getValues();
    for (var i = 0; i < POPRnumbers.length; i++) {
        var n = POPRnumbers[i][0];
        if (lowerLimit < n && n < upperLimit) {
            currentNumbers.push(n);
        }
    }
    if (currentNumbers.length == 0) {
        currentNumbers.push(lowerLimit);
    }
    currentNumbers.sort(function (a, b) { return b - a; });
    var _nextNumber = currentNumbers[0];
    _nextNumber = new Number(_nextNumber) + 1;
    nextNumber = _nextNumber.toString();
    nextNumber = '0000' + nextNumber;
    nextNumber = nextNumber.substr(-4);
    return nextNumber;
}
function getSupplier(supplierName) {
    aCD_WTSuppliers = getWTSup();
    var _sheet = SS.getSheetByName(LOCAL_SUPPLIERS_SHEET);
    var localSuppliers = _sheet.getRange(4, 1, (_sheet.getLastRow() - 3), _sheet.getLastColumn()).getValues();
    var allSuppliers = aCD_WTSuppliers.concat(localSuppliers);
    if (supplierName == '*all*') {
        return allSuppliers;
    }
    var I = -1;
    for (var i = 0; i < allSuppliers.length; i++) {
        if (allSuppliers[i][0] == supplierName) {
            I = i;
        }
    }
    var poSupplier = {
        name: '', contact: '', adr: '',
        tel: '', email: '',
        account: '', terms: '', admin: ''
    };
    if (I > -1) {
        poSupplier.name = supplierName;
        poSupplier.contact = allSuppliers[I][4];
        poSupplier.adr = '';
        poSupplier.tel = allSuppliers[I][5];
        poSupplier.email = allSuppliers[I][6];
        poSupplier.account = allSuppliers[I][2];
        poSupplier.terms = allSuppliers[I][3];
        poSupplier.admin = allSuppliers[I][1];
    }
    return poSupplier;
}
function POs4_updateMaterialsList(POs3) {
    var statusCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_STATUS) + 1), SH.getLastRow(), 1);
    var poNumCol = SH.getRange((MHRI + 2), (wholeList[MHRI].indexOf(H_PONUM) + 1), SH.getLastRow(), 1);
    var poCreatedCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_POCREATED) + 1), SH.getLastRow(), 1);
    var statusValues = statusCol.getValues();
    var poNumFormulas = poNumCol.getFormulas();
    var poCreatedValues = poCreatedCol.getValues();
    var poprLink = '=HYPERLINK(\"' + POs3.Url + '\",\"' + POs3.Num + '\")';
    SS.toast("PO created, adding links to materials list", "Step 5 of 5", 60);
    while (aRegenIndices.length) {
        var r = aRegenIndices.pop();
        statusValues[r][0] = '';
        poNumFormulas[(r - MHRI - 1)][0] = '';
        poCreatedValues[r][0] = '';
    }
    var _status = null;
    var _statusArray = getCentralDropDowns().ddStatusWT50;
    for (; _statusArray.length > 0;) {
        if (_statusArray[0].substr(0, 1) == REGEN_PREFIX) {
            _status = _statusArray[0];
        }
        _statusArray.shift();
    }
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
function generatePO(formObject) {
    sameDept = formObject.sameDept;
    delONBY = formObject.delivery;
    UI = SpreadsheetApp.getUi();
    var SHname = SH.getName();
    thisCell = SH.getActiveCell();
    if (SHname == MATERIALS_SHEET) {
        var thisRowIndex = SH.getActiveCell().getRow() - 1;
        if (thisRowIndex <= MHRI) {
            var wrongRow = UI.alert("Generate PO Error", "You have currently selected one of the header rows."
                + "\nPlease select a cell in a row from the materials list.", UI.ButtonSet.OK);
            return false;
        }
        var POs1 = POs1_getItems(thisRowIndex);
        var POs2 = POs2_handleIncompleteItems(POs1);
        if (iIndices.length > 250) {
            var tooManyItems = UI.alert("Too Many Items for One PO", "The maximum number of line items for a purchase order is 250.", UI.ButtonSet.OK);
        }
        if (POs2) {
            var BPRs = [];
            for (var noOfPRs = Math.ceil(iIndices.length / PR_ITEMS_LIMIT); noOfPRs > 0; noOfPRs--) {
                var _msgPR = (noOfPRs == 1) ? "1 PR with" : noOfPRs + " PRs from";
                var msgPR = (iIndices.length == 1) ? "Creating 1 PR with 1 valid item." : "Creating " + _msgPR + " " + iIndices.length + " valid items.";
                SS.toast(msgPR, "Step 4 of 5", 60);
                var PRs3 = PRs3_createPR();
                BPRs.push(PRs3);
                var PRs4 = PRs4_updateMaterialsList(PRs3);
            }
            SS.toast('Thank you for waiting', 'Done', 5);
        }
    }
    if (SHname != MATERIALS_SHEET) {
        SS.getSheetByName(MATERIALS_SHEET).activate();
        var uiResponse = UI.alert("Generate PO/PR", ". . . Changing Sheet to Materials List\nPlease click in a row you would like to generate a PO/PR for, then choose \"Generate PO/PR\" again.", UI.ButtonSet.OK);
    }
}
function POs1_getItems(thisRowIndex) {
    SS.toast("Compiling PO/PR Items, please wait", "Step 1 of 5", 5);
    wholeRange = SH.getRange(1, 1, SH.getLastRow(), SH.getLastColumn());
    wholeList = wholeRange.getValues();
    var incompleteItems = [];
    poSupplierName = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_SUPPLIER)];
    TYPE = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_TYPE)];
    poDelivery = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_ACTDEL)];
    var _poDelivery = '';
    if (poDelivery != '') {
        _poDelivery = new Date(poDelivery);
        var _poDelivery = Utilities.formatDate(new Date(poDelivery.setHours(8)), "GMT", "E, dd-MMM-yyyy");
    }
    poNumber = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_PONUM)];
    requestingDept = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_TEAM)];
    var thisLineStatus = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_STATUS)];
    branchPONumber = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_BRANCHPO)];
    if (wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_EMERGENCY)] == 'Yes') {
        isEmergency = true;
    }
    if (confirmGeneratePO(poSupplierName, _poDelivery, poNumber, requestingDept, thisLineStatus, isEmergency, branchPONumber)) {
        SS.toast("Sorting PO Items, please wait", "Step 3 of 5", 5);
        for (var i = MHRI + 1; i < wholeList.length; i++) {
            var _tempDelDate = new Date(wholeList[i][wholeList[MHRI].indexOf(H_ACTDEL)]);
            var tempDelDate = Utilities.formatDate(new Date(_tempDelDate.setHours(8)), "GMT", "E, dd-MMM-yyyy");
            var _thisBranchPO = wholeList[i][wholeList[MHRI].indexOf(H_BRANCHPO)];
            if (((wholeList[i][wholeList[MHRI].indexOf(H_PONUM)] == '' && wholeList[i][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) == ITEMREADY_PREFIX)
                || (wholeList[i][wholeList[MHRI].indexOf(H_PONUM)] == poNumber && wholeList[i][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) == REGEN_PREFIX))
                && (!sameDept || (sameDept && wholeList[i][wholeList[MHRI].indexOf(H_TEAM)] == requestingDept))
                && wholeList[i][wholeList[MHRI].indexOf(H_SUPPLIER)] == poSupplierName
                && wholeList[i][wholeList[MHRI].indexOf(H_TYPE)] == TYPE
                && ((!isEmergency && tempDelDate == _poDelivery) || (isEmergency))
                && (!isEmergency || (isEmergency && _thisBranchPO == branchPONumber))) {
                var _isEmergency = false;
                if (wholeList[i][wholeList[MHRI].indexOf(H_EMERGENCY)] == 'Yes') {
                    _isEmergency = true;
                }
                var _Qty = new Number(wholeList[i][wholeList[MHRI].indexOf(H_QTY)]);
                var _UoM = wholeList[i][wholeList[MHRI].indexOf(H_PUOM)];
                var _Factor = wholeList[i][wholeList[MHRI].indexOf(H_FACTOR)];
                var _BUoM = wholeList[i][wholeList[MHRI].indexOf(H_BUOM)];
                var _DESC = wholeList[i][wholeList[MHRI].indexOf(H_IDESC)];
                var _CODE = wholeList[i][wholeList[MHRI].indexOf(H_ICODE)];
                var _UNIT = new Number(wholeList[i][wholeList[MHRI].indexOf(H_UNIT)]);
                var _PDN = wholeList[i][wholeList[MHRI].indexOf(H_PDN)];
                var _PDC = wholeList[i][wholeList[MHRI].indexOf(H_PDC)];
                var _NOTEtemp = wholeList[i][wholeList[MHRI].indexOf(H_NOTES)];
                var _NOTE = '';
                if (delONBY == 'ON:' && !isEmergency) {
                    _NOTE = 'Please deliver ON: ' + _poDelivery;
                }
                if (delONBY == 'BY:' && !isEmergency) {
                    _NOTE = 'Please deliver from 2 days before: ' + _poDelivery;
                }
                if (isEmergency) {
                    _NOTE = 'Emergency Order: ' + branchPONumber;
                }
                if (_NOTE != '' && _NOTEtemp != '') {
                    _NOTE += ' ~ ' + _NOTEtemp;
                }
                if (_NOTE == '' && _NOTEtemp != '') {
                    _NOTE += _NOTEtemp;
                }
                if (TYPE == HIRE) {
                    var offHire = wholeList[i][wholeList[MHRI].indexOf(H_OFF)];
                    if (offHire != '') {
                        var _offHire = new Date(offHire);
                        offHire = Utilities.formatDate(new Date(_offHire.setHours(8)), "GMT", "E, dd-MMM-yyyy");
                        _NOTE += '{Hire from: ' + tempDelDate + ' to: ' + offHire + '}';
                    }
                    if (offHire == '' || _offHire < _tempDelDate) {
                        incompleteItems.push(i + 1);
                        continue;
                    }
                }
                if (_Qty > 0 && _UoM != '' && _Factor != '' && _BUoM != '' && _DESC != '' && (_CODE != '' || TYPE != CORELIST) && _UNIT > 0 && _PDN != '' && _PDC != ''
                    && ((_isEmergency && _thisBranchPO == branchPONumber) || (!_isEmergency && _thisBranchPO == ''))
                    && ((_isEmergency && tempDelDate == _poDelivery) || (!_isEmergency))) {
                    iQty.push(_Qty);
                    iUoM.push(_UoM);
                    iFactor.push(_Factor);
                    iBUoM.push(_BUoM);
                    if (TYPE == CORELIST || TYPE == COREEXTRA) {
                        iCode.push(_CODE);
                    }
                    iDesc.push(_DESC);
                    iUnit.push(_UNIT);
                    iPDN.push(_PDN);
                    iPDC.push(_PDC);
                    iNote.push(_NOTE);
                    iIndices.push(i);
                }
                else {
                    incompleteItems.push(i + 1);
                }
            }
        }
        return incompleteItems;
    }
    return null;
}
function confirmGeneratePO(poSupplier, poDelivery, poNumber, requestingDept, thisLineStatus, isEmergency, branchPONumber) {
    POorPR = 'PR';
    OrderOrRequest = 'Request';
    if (poNumber != '') {
        var poRegeneratePossible = true;
        var alreadyPOmsg = "This item has already been added to a Purchase " + OrderOrRequest + ", " + poNumber + ".";
        var alreadyPObtn = UI.ButtonSet.YES_NO;
        var poLink = SH.getRange(thisCell.getRow(), (wholeList[MHRI].indexOf(H_PONUM) + 1)).getFormula();
        var poID = poLink.substring(poLink.indexOf('/d/') + 3, poLink.indexOf('/edit'));
        for (var j = MHRI + 1; j < wholeList.length; j++) {
            if (wholeList[j][wholeList[MHRI].indexOf(H_PONUM)] == poNumber) {
                aRegenIndices.push(j + 1);
                if (wholeList[j][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) != REGEN_PREFIX) {
                    poRegeneratePossible = false;
                }
            }
        }
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
            var oldPO = DriveApp.getFileById(poID);
            oldPO.setName('TO DELETE >>> ' + oldPO.getName());
        }
    }
    if (isEmergency && branchPONumber == '') {
        var noBranchPO = UI.alert("Generate " + POorPR + " Error", "This line item is set as part of an emergency order, but no Branch PO has been entered.", UI.ButtonSet.OK);
        return false;
    }
    if (branchPONumber != '' && !isEmergency) {
        var branchPObutNotEmergency = UI.alert("Generate " + POorPR + " Error", "This line item has a Branch PO entered, but is not set as part of an emergency order.", UI.ButtonSet.OK);
        return false;
    }
    if (poSupplier == '') {
        var noSupplier = UI.alert("Generate " + POorPR + " Error", "No Supplier has been set for this line item. Please choose a valid supplier.", UI.ButtonSet.OK);
        return false;
    }
    if (poDelivery == '') {
        var noDeliveryDate = UI.alert("Generate " + POorPR + " Error", "No Delivery Date has been set for this line item. Please enter a valid date.", UI.ButtonSet.OK);
        return false;
    }
    if (thisLineStatus.substr(0, 1) != ITEMREADY_PREFIX && !poRegeneratePossible) {
        var wrongStatus = UI.alert("Generate " + POorPR + " Error", "This line item does not have the correct status to be processed.", UI.ButtonSet.OK);
        return false;
    }
    if ((poSupplier != '' || poRegeneratePossible) && poDelivery != '') {
        var msg = "Create a new " + POorPR + " for all to-be-ordered items from " + poSupplier + ", to be delivered " + delONBY + " " + poDelivery + "?";
        if (TYPE == HIRE) {
            msg = "Create a new " + POorPR + " for all to-be-hired items from " + poSupplier + ", to be delivered ON: " + poDelivery + "?"
                + "\n(Please note, hire " + POorPR + "s are always set to an ON delivery date, rather than BY)";
            delONBY = "ON:";
        }
        if (sameDept) {
            msg += "\n\n(Only items requested by \"" + requestingDept + "\" will be added to this " + POorPR + ")";
        }
        if (isEmergency) {
            msg = "Create a new " + POorPR + " for all emergency ordered items from " + poSupplier + ", on " + branchPONumber + "?";
            sameDept = false;
        }
        var confirmPO = UI.alert("Generate " + POorPR + " for " + poSupplier, msg, UI.ButtonSet.YES_NO);
        if (confirmPO == UI.Button.YES) {
            SS.toast("Collecting " + POorPR + " Items, please wait", "Step 2 of 5", 5);
            return true;
        }
        if (confirmPO == UI.Button.NO) {
            SS.toast(POorPR + " generation cancelled by user", POorPR + " Cancelled", 5);
            return false;
        }
    }
}
function POs2_handleIncompleteItems(POs1) {
    if (POs1.length == 0 && iIndices.length > 0) {
        return true;
    }
    if (POs1.length > 0 && iIndices.length == 0) {
        var nothingValidMessage = (POs1.length == 1) ? "Please check row " + POs1 + ", as it has missing information." : "Please check rows " + POs1 + ", as they have missing information.";
        var nothingValid = UI.alert("No Valid Items Found", nothingValidMessage, UI.ButtonSet.OK);
        var toastMsg1 = (POs1.length == 1) ? "Row to check: " + POs1 : "Rows to check: " + POs1;
        SS.toast(toastMsg1, "Incomplete Items", 60);
        return false;
    }
    if (POs1.length > 0 && isEmergency) {
        var emergencyMessage = (POs1.length == 1) ? "Please check row " + POs1 + ", as it has missing information." : "Please check rows " + POs1 + ", as they have missing information.";
        var emergencyAlert = UI.alert("Emergency Order Incomplete", emergencyMessage, UI.ButtonSet.OK);
        var toastMsg1e = (POs1.length == 1) ? "Row to check: " + POs1 : "Rows to check: " + POs1;
        SS.toast(toastMsg1e, "Incomplete Items", 60);
        return false;
    }
    var incompleteItemsTitle = '';
    var incompleteItemsMessage = '';
    if (POs1.length > 1) {
        incompleteItemsTitle = POs1.length + " items incomplete";
    }
    if (POs1.length > 1) {
        incompleteItemsMessage = POs1.length + " items out of " + (iIndices.length + POs1.length) + " are incomplete."
            + "\nYou can proceed with the " + POorPR + " without these items (OK) or CANCEL and fill in the missing information. The row numbers that need completing are: " + POs1
            + "\n(If you choose CANCEL a list of these row numbers will appear in the bottom right of the window as a reminder)";
    }
    if (POs1.length == 1) {
        incompleteItemsTitle = POs1.length + " item incomplete";
    }
    if (POs1.length == 1) {
        incompleteItemsMessage = POs1.length + " item out of " + (iIndices.length + POs1.length) + " is incomplete."
            + "\nYou can proceed with the " + POorPR + " without this item (OK) or CANCEL and fill in the missing information. The row number that needs completing is: " + POs1
            + "\n(If you choose CANCEL this row number will appear in the bottom right of the window as a reminder)";
    }
    if (POs1.length > 0) {
        var ignoreIncompleteItems = UI.alert(incompleteItemsTitle, incompleteItemsMessage, UI.ButtonSet.OK_CANCEL);
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
function PRs3_createPR() {
    var prMain = 'Import';
    var prMain_StartRow = 2;
    var newPRNumber = nextPOPRNumber();
    var newPRName = PROJ_NAME() + ' (' + PROJ_NUMBER() + ') BPR ' + newPRNumber;
    if (isEmergency) {
        newPRName += ' Emergency Order ' + branchPONumber;
    }
    var newPRFile = DriveApp.getFileById('1IbNvwMtwtwhje0_quRBEPeHdx7ohDLJ7NhOLZavww64')
        .makeCopy(newPRName, DriveApp.getFolderById(PR_FOLDER_ID()));
    newPRFile.setOwner(DRIVE_OWNER());
    var newPRUrl = newPRFile.getUrl();
    var newPRID = newPRUrl.substring(newPRUrl.indexOf('/d/') + 3, newPRUrl.indexOf('/edit'));
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(newPRID));
    var PR = SpreadsheetApp.getActiveSpreadsheet();
    var PRsh = PR.setActiveSheet(PR.getSheetByName(prMain));
    var li = iIndices.length;
    if (li > PR_ITEMS_LIMIT) {
        li = PR_ITEMS_LIMIT;
    }
    var retVal = { Num: newPRNumber, Url: newPRUrl, ID: newPRID, Name: newPRName, Indices: iIndices.slice(0, li) };
    var mainPRRange = PRsh.getRange(prMain_StartRow, 1, li, 27);
    var mainPRValues = mainPRRange.getValues();
    var daysToDelivery = dateDiffInDays(poDelivery, new Date());
    var bprPriority = 0;
    if (daysToDelivery < 8) {
        bprPriority = 1;
    }
    for (var i = 0; i < li; i++) {
        var thisIC = iCode.shift();
        mainPRValues[i][0] = thisIC;
        mainPRValues[i][1] = iDesc.shift();
        mainPRValues[i][2] = iQty.shift();
        mainPRValues[i][3] = iUoM.shift();
        mainPRValues[i][4] = iFactor.shift();
        mainPRValues[i][5] = iBUoM.shift();
        mainPRValues[i][6] = iUnit.shift();
        mainPRValues[i][7] = poDelivery;
        mainPRValues[i][8] = bprPriority;
        mainPRValues[i][9] = '';
        mainPRValues[i][10] = '';
        mainPRValues[i][11] = '';
        mainPRValues[i][12] = '';
        mainPRValues[i][13] = '';
        mainPRValues[i][14] = '';
        mainPRValues[i][15] = '';
        mainPRValues[i][16] = '';
        mainPRValues[i][17] = '';
        mainPRValues[i][18] = '';
        mainPRValues[i][19] = PROJ_NUMBER();
        mainPRValues[i][20] = iPDC.shift();
        mainPRValues[i][21] = '';
        mainPRValues[i][22] = '';
        mainPRValues[i][23] = iNote.shift();
        mainPRValues[i][24] = '';
        mainPRValues[i][25] = '';
        mainPRValues[i][26] = '';
    }
    mainPRRange.setValues(mainPRValues);
    SpreadsheetApp.setActiveSpreadsheet(SS);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SH).setActiveSelection(thisCell);
    return retVal;
}
function PRs4_updateMaterialsList(PRs3) {
    var statusCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_STATUS) + 1), SH.getLastRow(), 1);
    var poNumCol = SH.getRange((MHRI + 2), (wholeList[MHRI].indexOf(H_PONUM) + 1), SH.getLastRow(), 1);
    var poCreatedCol = SH.getRange(1, (wholeList[MHRI].indexOf(H_POCREATED) + 1), SH.getLastRow(), 1);
    var statusValues = statusCol.getValues();
    var poNumFormulas = poNumCol.getFormulas();
    var poCreatedValues = poCreatedCol.getValues();
    var poprLink = '=HYPERLINK(\"' + PRs3.Url + '\",\"' + PRs3.Num + '\")';
    SS.toast("PR " + PRs3.Num + " created, adding links to materials list", "Step 5 of 5", 60);
    while (aRegenIndices.length) {
        var r = aRegenIndices.pop();
        statusValues[r][0] = '';
        poNumFormulas[(r - MHRI - 1)][0] = '';
        poCreatedValues[r][0] = '';
    }
    var li = iIndices.length;
    if (li > PR_ITEMS_LIMIT) {
        li = PR_ITEMS_LIMIT;
    }
    var _status = null;
    var _statusArray = getCentralDropDowns().ddStatusBPR;
    for (; _statusArray.length > 0;) {
        if (_statusArray[0].substr(0, 1) == REGEN_PREFIX) {
            _status = _statusArray[0];
        }
        _statusArray.shift();
    }
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
function PR_Email(BPRs) {
    Logger.log(poSupplier);
    for (; BPRs.length > 0;) {
        var bpr = BPRs.shift();
        var emailSubject = bpr.Name;
        var _poDelivery = Utilities.formatDate(new Date(poDelivery.setHours(8)), "GMT", "E, dd-MMM-yyyy");
        var emailMessage = "Dear " + poSupplier.name
            + ",\n\nPlease find attached our Purchase Request (number " + bpr.Num + ") for the " + PROJ_NAME + " project."
            + "\nWe are requesting delivery " + delONBY + " " + _poDelivery;
        if (TYPE == COREEXTRA) {
            emailMessage += "\n\nPlease note that this Purchase Request includes items extra to the core list.";
        }
        emailMessage += "\n\nKind regards,\netc.";
        var bprFile = getAsExcel(bpr.ID);
        bprFile.setName(bpr.Name);
        MailApp.sendEmail(PRemailTo, emailSubject, emailMessage, {
            attachments: [bprFile],
            cc: PRemailCC
        });
    }
    return true;
}
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
        if (eValues[0][0] == 'A > Z') {
            sortDir = true;
        }
        if (eValues[0][0] == 'Z > A') {
            sortDir = false;
        }
        if (sortDir != null) {
            SH.getRange((svsSH_HR + 1), 1, SH.getLastRow(), SH.getLastColumn()).sort({ column: e.range.getColumn(), ascending: sortDir });
        }
        e.range.setValue('Sort?');
    }
}
function svsTester() {
    Logger.log(svsSH.getLastRow());
}
function getPRDataReadyForSVS() {
    var prData = [];
    var matData = matSH.getRange(1, 1, matSH.getLastRow(), matSH.getLastColumn()).getValues();
    var oldPRRange = svsSH.getRange((svsSH_HR), 1, (svsSH.getLastRow() - svsSH_HR), 5);
    var oldPRValues = oldPRRange.getValues();
    prData.push(oldPRValues[0]);
    if (oldPRValues.length > 0) {
        oldPRRange.clearContent();
    }
    for (var m = MHRI + 1; m < matData.length; m++) {
        if (matData[m][matData[MHRI].indexOf(H_PONUM)] < 1000
            && matData[m][matData[MHRI].indexOf(H_PONUM)] > 0
            && matData[m][matData[MHRI].indexOf(H_STATUS)].substr(0, 1) == EXPECTINGSVS_PREFIX) {
            var _prData = [];
            _prData.push(matData[m][matData[MHRI].indexOf(H_PONUM)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_LINEID)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_IDESC)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_QTY)]);
            _prData.push(matData[m][matData[MHRI].indexOf(H_ACTDEL)]);
            prData.push(_prData);
        }
    }
    if (prData) {
        var thisPRRange = svsSH.getRange(svsSH_HR, 1, prData.length, prData[0].length);
        thisPRRange.setValues(prData);
    }
    return prData;
}
function confirmMatchedItems() {
    SS.toast("Reviewing PR / SVS Data matches", "SVS Confirm 1 of 4", 30);
    var currentSVSRange = svsSH.getRange(svsSH_HR, 1, (svsSH.getLastRow() - svsSH_HR), svsSH.getLastColumn());
    var currentSVSValues = currentSVSRange.getValues();
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
    svsToBeMatched.push(currentSVSValues[0]);
    SS.toast("Getting SVS PDF files", "SVS Confirm 2 of 4", 30);
    for (var s = 1; s < currentSVSValues.length; s++) {
        var svsMatchID = '';
        if (currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_SUPPLIER)] != ''
            || currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_NUM)] != ''
            || currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_LINE)] != '') {
            svsMatchID = currentSVSValues[s][currentSVSValues[0].indexOf(svsMAT_LINE)];
        }
        for (var t = 0; t < lineIDValues.length; t++) {
            var thisMaterialLineID = lineIDValues[t][0];
            if (svsMatchID == thisMaterialLineID && svsMatchID != '') {
                orderStatusValues[t][0] = statusSVSReceived;
                var wtpoNum = currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_NUM)];
                var wtsvsUrl = 'https://drive.google.com/open?id=0B6duJkKLBboAZVFsMW45eHg1Qms';
                branchPOFormulas[t][0] = '=HYPERLINK(\"' + wtsvsUrl + '\",\"' + wtpoNum + '\")';
                branchPOValues[t][0] = '=HYPERLINK(\"' + wtsvsUrl + '\",\"' + wtpoNum + '\")';
                branchLineValues[t][0] = currentSVSValues[s][currentSVSValues[0].indexOf(svsSVS_WTPO_LINE)];
            }
        }
    }
    SS.toast("Writing PR / SVS Data matches", "SVS Confirm 3 of 4", 30);
    orderStatusCol.setValues(orderStatusValues);
    branchPOCol.setFormulas(branchPOFormulas);
    branchLineCol.setValues(branchLineValues);
    branchSupplierCol.setValues(branchSupplierValues);
    SS.toast("Refreshing SVS Matcher", "SVS Confirm 4 of 4", 5);
    currentSVSRange.clearContent().clearNote();
    var refreshedSVSRange = svsSH.getRange(svsSH_HR, 1, svsToBeMatched.length, svsToBeMatched[0].length);
    refreshedSVSRange.setValues(svsToBeMatched);
    var thisPRData = getPRDataReadyForSVS();
}
function getPDFSVS() {
    SS.toast("Searching for valid SVS files", "SVS Process 1 of 4", 30);
    var SVSFolder = DriveApp.getFolderById(SVS_FOLDER_ID().toDo);
    var SVSFiles = SVSFolder.getFiles();
    var svsPDFPR = [];
    var svsPDFURL = [];
    while (SVSFiles.hasNext()) {
        var svsFile = SVSFiles.next();
        var fileName = svsFile.getName();
        var svsProjNum = fileName.substr(4, 3).toString();
        var svsPRNum = fileName.substr(8, 4).toString();
        if (svsProjNum == PROJ_NUMBER()) {
            if (svsFile.getMimeType() == 'application/pdf') {
                svsPDFURL.push(svsFile.getUrl());
                svsPDFPR.push(svsPRNum);
            }
        }
    }
}
