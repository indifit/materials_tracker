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
} // end fn:clearCells
//# sourceMappingURL=CoreList.js.map
