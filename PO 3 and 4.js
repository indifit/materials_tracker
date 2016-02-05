/* * * * * * *
* STAGE 3 of PO Generation Materials - create the PO
*
* * * * * * */
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

    var newPOFile = DriveApp.getFileById(PO_TEMPLATE_ID).makeCopy(newPOName + ' ~ ' + poSupplierName, DriveApp.getFolderById(PO_FOLDER_ID));
    newPOFile.setOwner(DRIVE_OWNER);

    var newPOUrl = newPOFile.getUrl();

    var newPOID = newPOUrl.substring(newPOUrl.indexOf('/d/') + 3, newPOUrl.indexOf('/edit'));

    var retVal = { Num: newPONumber, Url: newPOUrl, ID: newPOID, Name: newPOName };

    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(newPOID));
    var PO = SpreadsheetApp.getActiveSpreadsheet();
    var POsh = SpreadsheetApp.getActiveSheet();

    var i = iIndices.length;
    var appSubs = [];

    // if i>10, then calculate how many appendicies, a, are needed
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
}

function nextPOPRNumber() {
    // if a PO number already exists (i.e. regeneration) just return the same PO number
    if (poNumber != '') {
        return poNumber;
    }

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
        if (lowerLimit < n && n < upperLimit) {
            currentNumbers.push(n);
        }
    }

    if (currentNumbers.length == 0) {
        currentNumbers.push(lowerLimit);
    }

    // sort the list of current PO numbers to get the highest one
    currentNumbers.sort(function (a, b) {
        return b - a;
    });
    var _nextNumber = currentNumbers[0];

    _nextNumber = new Number(_nextNumber) + 1;
    nextNumber = _nextNumber.toString();
    nextNumber = '0000' + nextNumber;
    nextNumber = nextNumber.substr(-4);

    return nextNumber;
}

function getSupplier(supplierName) {
    // set up placeholders
    aCD_WTSuppliers = getWTSup();

    var _sheet = SS.getSheetByName(LOCAL_SUPPLIERS_SHEET);
    var localSuppliers = _sheet.getRange(4, 1, (_sheet.getLastRow() - 3), _sheet.getLastColumn()).getValues();

    var allSuppliers = aCD_WTSuppliers.concat(localSuppliers);

    if (supplierName == '*all*') {
        return allSuppliers;
    }

    // the Index for the "found" supplier
    var I = -1;

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
}

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
//# sourceMappingURL=PO 3 and 4.js.map
