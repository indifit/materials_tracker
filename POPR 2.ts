/* * * * * * *
* STAGE 2 of PO Generation Materials - handle any incomplete items
*   Alert if any items have missing information
*     and ask whether to proceed anyway or cancel and correct
* * * * * * */

function POs2_handleIncompleteItems(POs1) {

    // if there are no incomplete items and there is at least one valid item, return true
    if (POs1.length == 0 && iIndices.length > 0) { return true; }

    // if there are only incomplete items but no valid ones, alert and then return false
    if (POs1.length > 0 && iIndices.length == 0) {
        var nothingValidMessage = (POs1.length == 1) ? "Please check row " + POs1 + ", as it has missing information." : "Please check rows " + POs1 + ", as they have missing information.";
        var nothingValid = UI.alert("No Valid Items Found", nothingValidMessage, UI.ButtonSet.OK);
        var toastMsg1 = (POs1.length == 1) ? "Row to check: " + POs1 : "Rows to check: " + POs1;
        SS.toast(toastMsg1, "Incomplete Items", 60);
        return false;
    }

    // if there are any incomplete items and this is an emergency order, alert and then return false
    if (POs1.length > 0 && isEmergency) {
        var emergencyMessage = (POs1.length == 1) ? "Please check row " + POs1 + ", as it has missing information." : "Please check rows " + POs1 + ", as they have missing information.";
        var emergencyAlert = UI.alert("Emergency Order Incomplete", emergencyMessage, UI.ButtonSet.OK);
        var toastMsg1e = (POs1.length == 1) ? "Row to check: " + POs1 : "Rows to check: " + POs1;
        SS.toast(toastMsg1e, "Incomplete Items", 60);
        return false;
    }

    var incompleteItemsTitle = '';
    var incompleteItemsMessage = '';

    if (POs1.length > 1) { incompleteItemsTitle = POs1.length + " items incomplete"; }
    if (POs1.length > 1) {
        incompleteItemsMessage = POs1.length + " items out of " + (iIndices.length + POs1.length) + " are incomplete."
        + "\nYou can proceed with the " + POorPR + " without these items (OK) or CANCEL and fill in the missing information. The row numbers that need completing are: " + POs1
        + "\n(If you choose CANCEL a list of these row numbers will appear in the bottom right of the window as a reminder)";
    }

    if (POs1.length == 1) { incompleteItemsTitle = POs1.length + " item incomplete"; }
    if (POs1.length == 1) {
        incompleteItemsMessage = POs1.length + " item out of " + (iIndices.length + POs1.length) + " is incomplete."
        + "\nYou can proceed with the " + POorPR + " without this item (OK) or CANCEL and fill in the missing information. The row number that needs completing is: " + POs1
        + "\n(If you choose CANCEL this row number will appear in the bottom right of the window as a reminder)";
    }

    if (POs1.length > 0) {
        var ignoreIncompleteItems = UI.alert(incompleteItemsTitle, incompleteItemsMessage, UI.ButtonSet.OK_CANCEL);
        // Process the user's response.
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