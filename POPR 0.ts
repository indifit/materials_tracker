// single values for PO entries from materials list of thisRow
var poSupplierName = '';
var poDelivery = '';
var poNumber = '';
var requestingDept = '';
var sameDept = '';
var delONBY = '';

// multi dimensional array holder for the whole materials sheet
var wholeRange = '';
var wholeList = [];
// multi dimensional array holder for the wt suppliers import
var aCD_WTSuppliers = [];

// arrays for PO entries from materials list
var iQty = [];
var iUoM = [];
var iFactor = [];
var iBUoM = [];
var iCode = []; // item code used separately for PRs
var iDesc = []; // item description is combined with [item code] for Materials and Hire. Also combined with on delivery / off delivery for hire items
var iUnit = [];
var iPDN = []; // project dimension name, not actually on PO, but place in hidden column
var iPDC = []; // project dimension code, used for BPR
var iIndices = []; // indices numbers for line items to have a PO added
var iNote = [];

var aRegenIndices = []; // array of indices to remove regen links

var isEmergency = false;
var branchPONumber = '';

/*
 * ciriteria for a PO:
 * same... supplier, delivery date and type (possibly limit to same department)
 *
*/

function generatePO(formObject) {

    sameDept = formObject.sameDept;
    delONBY = formObject.delivery;

    UI = SpreadsheetApp.getUi();
    var SHname = SH.getName();
    thisCell = SH.getActiveCell();

    if (SHname == MATERIALS_SHEET) {

        // thisRowIndex is the current row -1 to allow for zero index in arrays
        var thisRowIndex = SH.getActiveCell().getRow() - 1;

        // if a cell is selected in a header row, not an item row
        if (thisRowIndex <= MHRI) {
            var wrongRow = UI.alert("Generate PO Error",
                "You have currently selected one of the header rows."
                + "\nPlease select a cell in a row from the materials list.",
                UI.ButtonSet.OK);
            return false;
        }

        // * * * * stage 1, get the materials * * * 
        var POs1 = POs1_getItems(thisRowIndex); // returns an array of row numbers of any incomplete items

        // * * * * stage 2, handle any incomplete items * * *
        var POs2 = POs2_handleIncompleteItems(POs1); // returns true or false to proceed

        // if there are more than 250 items and a PO is needed
        if (iIndices.length > 250) {
            var tooManyItems = UI.alert("Too Many Items for One PO",
                "The maximum number of line items for a purchase order is 250.",// The first 250 valid items will be added to this PO then the remainder will be added to a new PO"
                UI.ButtonSet.OK);
        }
        // ********* this needs doing    

        /*    // if there are less than 250 items and a PO is needed
            if (POs2 && iIndices.length <= 250 && TYPE!=CORELIST && TYPE!=COREEXTRA) {
              var POs2_msg = (iIndices.length == 1) ? "Creating PO with 1 valid item.":"Creating PO with "+iIndices.length+" valid items.";
              SS.toast(POs2_msg,"Step 4 of 5",60);
              var POs3 = POs3_createPO();
              // clean up and add the PO links
              var POs4 = POs4_updateMaterialsList(POs3);
              if (POs4){SS.toast("PO "+POs3.Num+" has been successfully created","Done!",5);}
            }
        */
        // if a BPR is needed
        if (POs2) {

            // setup holders for the completed message

            // an array to hold all the produced BPR data
            var BPRs = [];

            // use a loop to create the PRs
            for (var noOfPRs = Math.ceil(iIndices.length / PR_ITEMS_LIMIT); noOfPRs > 0; noOfPRs--) {

                var _msgPR = (noOfPRs == 1) ? "1 PR with" : noOfPRs + " PRs from";
                var msgPR = (iIndices.length == 1) ? "Creating 1 PR with 1 valid item." : "Creating " + _msgPR + " " + iIndices.length + " valid items.";
                SS.toast(msgPR, "Step 4 of 5", 60);

                var PRs3 = PRs3_createPR();
                BPRs.push(PRs3);
                // clean up and add the PO links
                var PRs4 = PRs4_updateMaterialsList(PRs3);

            } // end for loop
            SS.toast('Thank you for waiting', 'Done', 5);

            /*      // ask if newly created BPRs should be emailed now
                  var bprTitle = '';
                  var bprMessage = '';
  
                  var noBPRs = BPRs.length;
      
                  if (noBPRs==1) { bprTitle = noBPRs+" Purchase Request Created"; }
                  if (noBPRs==1) { bprMessage = noBPRs+" has been successfully created. Would you like to email it to "+poSupplier.name+" now?"; }
      
                  if (noBPRs>1) { bprTitle = noBPRs+" Purchase Requests Created"; }
                  if (noBPRs>1) { bprMessage = noBPRs+" have been successfully created. Would you like to email them to "+poSupplier.name+" now?"; }
      

                  if (noBPRs>0) {

                    var bprEmail = UI.alert(bprTitle, bprMessage, UI.ButtonSet.YES_NO);

                    // Process the user's response.
                    if (bprEmail == UI.Button.YES) {
                      var bprEmailToast = (noBPRs == 1) ? "Sending "+noBPRs+" PR to "+poSupplier.name:"Sending "+noBPRs+" PRs to "+poSupplier.name;
                      SS.toast(bprEmailToast,"Emailing PR",60);
          
                      // call the PR_Email function and wait for a true return
                      if (PR_Email(BPRs)){
                        var bprEmailedToast = (noBPRs == 1) ? noBPRs+" PR sent to "+poSupplier.name:noBPRs+" PRs sent to "+poSupplier.name;
                        SS.toast(bprEmailedToast,"Done!",5);
                      }
          
                    } // end if user said yes to sending email(s)
                  } // end if BPRs were created
            */
        } // end if BPR is needed

    } // end if correct sheet

    // if not on the materials sheet
    if (SHname != MATERIALS_SHEET) {
        SS.getSheetByName(MATERIALS_SHEET).activate();
        var uiResponse = UI.alert("Generate PO/PR",
            ". . . Changing Sheet to Materials List\nPlease click in a row you would like to generate a PO/PR for, then choose \"Generate PO/PR\" again.",
            UI.ButtonSet.OK);
    } // end if wrong sheet

} // end fn:generatePO_materials