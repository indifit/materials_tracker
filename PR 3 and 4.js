/* * * * * * *
* STAGE 3 of PR Generation Materials - create the PR
*
* * * * * * */
function PRs3_createPR() {
    var prProjectNameCell = 'C2';
    var prPRNumberCell = 'K1';
    var prProjectNumberCell = 'K2';
    var emergencyHeaderCell = 'D1';

    var prMain = 'To be completed';
    var prMain_StartRow = 4;

    var newPRNumber = nextPOPRNumber();
    var newPRName = PROJ_NAME() + ' (' + PROJ_NUMBER() + ') BPR ' + newPRNumber;
    if (isEmergency) {
        newPRName += ' Emergency Order ' + branchPONumber;
    }

    var newPRFile = DriveApp.getFileById(PR_TEMPLATE_ID()).makeCopy(newPRName, DriveApp.getFolderById(PR_FOLDER_ID()));

    newPRFile.setOwner(DRIVE_OWNER());

    var newPRUrl = newPRFile.getUrl();

    var newPRID = newPRUrl.substring(newPRUrl.indexOf('/d/') + 3, newPRUrl.indexOf('/edit'));

    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(newPRID));
    var PR = SpreadsheetApp.getActiveSpreadsheet();
    var PRsh = PR.setActiveSheet(PR.getSheetByName(prMain));

    // write the single values to the new PR
    if (isEmergency) {
        PR.getRange(emergencyHeaderCell).setValue('Emergency Order ' + branchPONumber);
    }
    PR.getRange(prProjectNameCell).setValue(PROJ_NAME());
    PR.getRange(prProjectNumberCell).setValue(PROJ_NUMBER());
    PR.getRange(prPRNumberCell).setValue(newPRNumber);

    // cap the number of line items to the PR Items limit
    var li = iIndices.length;
    if (li > PR_ITEMS_LIMIT) {
        li = PR_ITEMS_LIMIT;
    }

    // write the retVal value, but add the indices that this PR refers to
    var retVal = { Num: newPRNumber, Url: newPRUrl, ID: newPRID, Name: newPRName, Indices: iIndices.slice(0, li) };

    // get the body area of the PR page
    var mainPRRange = PRsh.getRange(prMain_StartRow, 1, li, 11);
    var mainPRValues = mainPRRange.getValues();

    // get the number of days between today and the delivery date
    var daysToDelivery = dateDiffInDays(poDelivery, new Date());

    var bprPriority = 'N';
    if (daysToDelivery < 8) {
        bprPriority = 'Y';
    }

    for (var i = 0; i < li; i++) {
        var thisIC = iCode.shift();
        mainPRValues[i][0] = thisIC;
        mainPRValues[i][1] = iDesc.shift();
        mainPRValues[i][2] = iUnit.shift();
        mainPRValues[i][3] = iUoM.shift();
        mainPRValues[i][4] = iFactor.shift();
        mainPRValues[i][5] = iBUoM.shift();
        mainPRValues[i][6] = iQty.shift();
        mainPRValues[i][7] = iPDC.shift();
        mainPRValues[i][8] = poDelivery;
        mainPRValues[i][9] = bprPriority;
        mainPRValues[i][10] = iNote.shift();
    }

    mainPRRange.setValues(mainPRValues);

    // return to original sheet
    SpreadsheetApp.setActiveSpreadsheet(SS);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SH).setActiveSelection(thisCell);

    // return the details of the new PR to the main script
    return retVal;
}

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

    while (aRegenIndices.length) {
        var r = aRegenIndices.pop();
        statusValues[r][0] = '';
        poNumFormulas[(r - MHRI - 1)][0] = '';
        poCreatedValues[r][0] = '';
    }

    // cap the number of line items to the PR Items limit
    var li = iIndices.length;
    if (li > PR_ITEMS_LIMIT) {
        li = PR_ITEMS_LIMIT;
    }

    var _status = null;
    var _statusArray = getCentralDropDowns().ddStatusBPR;
    for (; _statusArray.length > 0;) {
        // get the status dropdown value whose prefix matches the regen prefix
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
//# sourceMappingURL=PR 3 and 4.js.map
