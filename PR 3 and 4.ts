/* * * * * * *
* STAGE 3 of PR Generation Materials - create the PR
*  
* * * * * * */

function PRs3_createPR() {

  var prMain = 'Import';
  var prMain_StartRow = 2;
  
  var newPRNumber = nextPOPRNumber();
  var newPRName = PROJ_NAME() +' ('+ PROJ_NUMBER() +') BPR '+ newPRNumber;
  if (isEmergency){newPRName += ' Emergency Order '+branchPONumber;}
  
  var newPRFile = DriveApp.getFileById('1IbNvwMtwtwhje0_quRBEPeHdx7ohDLJ7NhOLZavww64')
                         .makeCopy(newPRName,
                                   DriveApp.getFolderById(PR_FOLDER_ID()));
  
  newPRFile.setOwner(DRIVE_OWNER());

  var newPRUrl = newPRFile.getUrl();
    
  var newPRID = newPRUrl.substring(newPRUrl.indexOf('/d/')+3,newPRUrl.indexOf('/edit'));

  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(newPRID));
  var PR = SpreadsheetApp.getActiveSpreadsheet();
  var PRsh = PR.setActiveSheet(PR.getSheetByName(prMain));
  
  // cap the number of line items to the PR Items limit
  var li = iIndices.length;
  if (li>PR_ITEMS_LIMIT){li=PR_ITEMS_LIMIT;}
  
  // write the retVal value, but add the indices that this PR refers to
  var retVal = {Num:newPRNumber, Url:newPRUrl, ID:newPRID, Name:newPRName, Indices:iIndices.slice(0,li)};
  
  // get the body area of the PR page
  var mainPRRange = PRsh.getRange(prMain_StartRow, 1, li, 27);
  var mainPRValues = mainPRRange.getValues();
  
  // get the number of days between today and the delivery date
  var daysToDelivery = dateDiffInDays(poDelivery, new Date()); 

  var bprPriority = 0;
  if (daysToDelivery<8){bprPriority = 1;}
  
  // fill the main PR page with line items - use shift to start from the top of the list
  // [row][column] but zero indexed not 1 (therefore -1)
  for(var i=0;i<li;i++) {
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