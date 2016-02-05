function PR_Email(BPRs) {

    Logger.log(poSupplier);

    for (; BPRs.length > 0;) {

        var bpr = BPRs.shift();

        var emailSubject = bpr.Name;

        // the date is set to 8am to remove a problem of GMT dates being shown as the day before when the "now" date is BST
        var _poDelivery = Utilities.formatDate(new Date(poDelivery.setHours(8)), "GMT", "E, dd-MMM-yyyy");

        var emailMessage = "Dear " + poSupplier.name
            + ",\n\nPlease find attached our Purchase Request (number " + bpr.Num + ") for the " + PROJ_NAME + " project."
            + "\nWe are requesting delivery " + delONBY + " " + _poDelivery;
        if (TYPE == COREEXTRA) { emailMessage += "\n\nPlease note that this Purchase Request includes items extra to the core list."; }
        emailMessage += "\n\nKind regards,\netc."

    // get the active PO Google Sheets file ID

    var bprFile = getAsExcel(bpr.ID);//DriveApp.getFileById(bpr.ID);
        bprFile.setName(bpr.Name)

    // send the email as the current user

    MailApp.sendEmail(PRemailTo,
            emailSubject,
            emailMessage,
            {
                attachments: [bprFile],//.getAs(MimeType.MICROSOFT_EXCEL)],
                cc: PRemailCC
            }
            );

    } // end for loop

    return true;

} // end fn:PR_Email

// this uses the advance API Drive service, which was enabled in both the Resources>Advanced Google services... menu and the developer api console
function getAsExcel(spreadsheetId) {
    var file = Drive.Files.get(spreadsheetId);
    var url = file.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });
    return response.getBlob();
}