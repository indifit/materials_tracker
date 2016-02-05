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
var svsMAT_DEL = 'Requested Delivery Date'

var svsSVS_SUPPLIER = 'SVS Supplier';
var svsSVS_WTPO_NUM = 'SVS WT PO Number';
var svsSVS_WTPO_LINE = 'SVS WT PO Line';


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