var GoodReceiving;
(function (GoodReceiving) {
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
    function setupGoodsReceiver() {
    }
    var MATLIST = null;
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
})(GoodReceiving || (GoodReceiving = {}));
