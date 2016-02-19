var CoreList;
(function (CoreList) {
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
})(CoreList || (CoreList = {}));
