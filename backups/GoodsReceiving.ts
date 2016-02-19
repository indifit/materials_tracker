module GoodReceiving {
    /*
    * some functionality is shared with the core list 
    * particularly dropdown filtering and display.
    * however, to maintain independance, functions are duplicated
    */

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

    var MATLIST = null;
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
    
}
