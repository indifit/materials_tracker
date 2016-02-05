/* * * * * * *
* STAGE 1 of PO Generation Materials - get materials and validate items
*   get the whole sheet of items
*   extract only matching Supplier and Delivery date
*   Alert if any items are blank or else
*     write item arrays for ready for PO
* * * * * * */
function POs1_getItems(thisRowIndex) {
    SS.toast("Compiling PO/PR Items, please wait", "Step 1 of 5", 5);

    // import materials list
    wholeRange = SH.getRange(1, 1, SH.getLastRow(), SH.getLastColumn());

    wholeList = wholeRange.getValues(); // an array of rows, each an array of columns wholeList[r-1][c-1]

    var incompleteItems = [];

    // single values for PO entries from materials list of thisRow
    poSupplierName = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_SUPPLIER)];
    TYPE = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_TYPE)];

    poDelivery = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_ACTDEL)];
    var _poDelivery = '';
    if (poDelivery != '') {
        _poDelivery = new Date(poDelivery);

        // the date is set to 8am to remove a problem of GMT dates being shown as the day before when the "now" date is BST
        var _poDelivery = Utilities.formatDate(new Date(poDelivery.setHours(8)), "GMT", "E, dd-MMM-yyyy");
    }
    poNumber = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_PONUM)];
    requestingDept = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_TEAM)];

    var thisLineStatus = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_STATUS)];

    branchPONumber = wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_BRANCHPO)];

    if (wholeList[thisRowIndex][wholeList[MHRI].indexOf(H_EMERGENCY)] == 'Yes') {
        isEmergency = true;
    }

    // if the PO Generation has been confirmed based on this line item, collect other matching items
    if (confirmGeneratePO(poSupplierName, _poDelivery, poNumber, requestingDept, thisLineStatus, isEmergency, branchPONumber)) {
        SS.toast("Sorting PO Items, please wait", "Step 3 of 5", 5);

        for (var i = MHRI + 1; i < wholeList.length; i++) {
            var _tempDelDate = new Date(wholeList[i][wholeList[MHRI].indexOf(H_ACTDEL)]);
            var tempDelDate = Utilities.formatDate(new Date(_tempDelDate.setHours(8)), "GMT", "E, dd-MMM-yyyy");

            var _thisBranchPO = wholeList[i][wholeList[MHRI].indexOf(H_BRANCHPO)];

            if (((wholeList[i][wholeList[MHRI].indexOf(H_PONUM)] == '' && wholeList[i][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) == ITEMREADY_PREFIX) || (wholeList[i][wholeList[MHRI].indexOf(H_PONUM)] == poNumber && wholeList[i][wholeList[MHRI].indexOf(H_STATUS)].substr(0, 1) == REGEN_PREFIX)) && (!sameDept || (sameDept && wholeList[i][wholeList[MHRI].indexOf(H_TEAM)] == requestingDept)) && wholeList[i][wholeList[MHRI].indexOf(H_SUPPLIER)] == poSupplierName && wholeList[i][wholeList[MHRI].indexOf(H_TYPE)] == TYPE && ((!isEmergency && tempDelDate == _poDelivery) || (isEmergency)) && (!isEmergency || (isEmergency && _thisBranchPO == branchPONumber))) {
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

                //        if (_CODE!='' && TYPE!=CORELIST && TYPE!=COREEXTRA) {_DESC += ' ['+_CODE+']';} // quick append of code to description if this is a PO, not PR
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

                // if this is a PR for a hire item, get the off hire date
                if (TYPE == HIRE) {
                    var offHire = wholeList[i][wholeList[MHRI].indexOf(H_OFF)];

                    if (offHire != '') {
                        var _offHire = new Date(offHire);

                        // the date is set to 8am to remove a problem of GMT dates being shown as the day before when the "now" date is BST
                        offHire = Utilities.formatDate(new Date(_offHire.setHours(8)), "GMT", "E, dd-MMM-yyyy");
                        _NOTE += '{Hire from: ' + tempDelDate + ' to: ' + offHire + '}'; // update the description with item specific on hire / off hire details.
                    }

                    if (offHire == '' || _offHire < _tempDelDate) {
                        incompleteItems.push(i + 1); // ,register an incomplete item
                        continue;
                    }
                }

                // if all elements are present, push into array
                if (_Qty > 0 && _UoM != '' && _Factor != '' && _BUoM != '' && _DESC != '' && (_CODE != '' || TYPE != CORELIST) && _UNIT > 0 && _PDN != '' && _PDC != '' && ((_isEmergency && _thisBranchPO == branchPONumber) || (!_isEmergency && _thisBranchPO == '')) && ((_isEmergency && tempDelDate == _poDelivery) || (!_isEmergency))) {
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
                } else {
                    incompleteItems.push(i + 1);
                }
            }
        }

        return incompleteItems;
    }

    return null;
}

/*
*
*/
function confirmGeneratePO(poSupplier, poDelivery, poNumber, requestingDept, thisLineStatus, isEmergency, branchPONumber) {
    POorPR = 'PR';
    OrderOrRequest = 'Request';

    // if a PO number is already present
    if (poNumber != '') {
        var poRegeneratePossible = true;
        var alreadyPOmsg = "This item has already been added to a Purchase " + OrderOrRequest + ", " + poNumber + ".";
        var alreadyPObtn = UI.ButtonSet.YES_NO;
        var poLink = SH.getRange(thisCell.getRow(), (wholeList[MHRI].indexOf(H_PONUM) + 1)).getFormula();
        var poID = poLink.substring(poLink.indexOf('/d/') + 3, poLink.indexOf('/edit'));

        for (var j = MHRI + 1; j < wholeList.length; j++) {
            if (wholeList[j][wholeList[MHRI].indexOf(H_PONUM)] == poNumber) {
                aRegenIndices.push(j + 1); // write the indices for the all currently assigned POs
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
            alreadyPOmsg += "\nHowever, it is possible to regenerate this " + POorPR + " to include all valid items. Do you want to regenerate this " + POorPR + "?";
        }

        var alreadyPO = UI.alert("Generate " + POorPR + " Error", alreadyPOmsg, alreadyPObtn);

        if (!poRegeneratePossible || (alreadyPO == UI.Button.NO)) {
            SS.toast(POorPR + " generation cancelled", POorPR + " Cancelled", 5);
            return false;
        }

        if (alreadyPO == UI.Button.YES) {
            SS.toast("Deleting existing " + POorPR + " and removing links, please wait", "Deleting " + POorPR + " " + poNumber, 60);

            // mark for deletion
            var oldPO = DriveApp.getFileById(poID);
            oldPO.setName('TO DELETE >>> ' + oldPO.getName());
        }
    }

    // if this is an emergency order, but no Branch PO number is present
    if (isEmergency && branchPONumber == '') {
        var noBranchPO = UI.alert("Generate " + POorPR + " Error", "This line item is set as part of an emergency order, but no Branch PO has been entered.", UI.ButtonSet.OK);

        // ***** possibly setActiveSelection to the cell that needs to be completed
        return false;
    }

    // if a Branch PO number is present, but this item is not set to be an emergency order
    if (branchPONumber != '' && !isEmergency) {
        var branchPObutNotEmergency = UI.alert("Generate " + POorPR + " Error", "This line item has a Branch PO entered, but is not set as part of an emergency order.", UI.ButtonSet.OK);

        // ***** possibly setActiveSelection to the cell that needs to be completed
        return false;
    }

    // if the supplier is empty
    if (poSupplier == '') {
        var noSupplier = UI.alert("Generate " + POorPR + " Error", "No Supplier has been set for this line item. Please choose a valid supplier.", UI.ButtonSet.OK);

        // ***** possibly setActiveSelection to the cell that needs to be completed
        return false;
    }

    // if the delivery date is empty
    if (poDelivery == '') {
        var noDeliveryDate = UI.alert("Generate " + POorPR + " Error", "No Delivery Date has been set for this line item. Please enter a valid date.", UI.ButtonSet.OK);

        // ***** possibly setActiveSelection to the cell that needs to be completed
        return false;
    }

    // if the status is not valid
    if (thisLineStatus.substr(0, 1) != ITEMREADY_PREFIX && !poRegeneratePossible) {
        var wrongStatus = UI.alert("Generate " + POorPR + " Error", "This line item does not have the correct status to be processed.", UI.ButtonSet.OK);

        // ***** possibly setActiveSelection to the cell that needs to be completed
        return false;
    }

    // if there is a Suppiler and Delivery date entered for the item in this row
    if ((poSupplier != '' || poRegeneratePossible) && poDelivery != '') {
        var msg = "Create a new " + POorPR + " for all to-be-ordered items from " + poSupplier + ", to be delivered " + delONBY + " " + poDelivery + "?";

        if (TYPE == HIRE) {
            msg = "Create a new " + POorPR + " for all to-be-hired items from " + poSupplier + ", to be delivered ON: " + poDelivery + "?" + "\n(Please note, hire " + POorPR + "s are always set to an ON delivery date, rather than BY)";
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

        // Process the user's response.
        if (confirmPO == UI.Button.YES) {
            SS.toast("Collecting " + POorPR + " Items, please wait", "Step 2 of 5", 5);
            return true;
        }
        if (confirmPO == UI.Button.NO) {
            SS.toast(POorPR + " generation cancelled by user", POorPR + " Cancelled", 5);
            return false;
        }
    }
} // end fn:confirmGeneratePO
//# sourceMappingURL=POPR 1.js.map
