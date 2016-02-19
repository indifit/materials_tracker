var UI = null;
var Code;
(function (Code) {
    function onOpen(e) {
        UI = SpreadsheetApp.getUi();
        var statusUpdateSubMenu = UI.createMenu('Update Order Status')
            .addItem('...to Sent', 'updateStatustoSent');
        UI.createMenu('Purchasing Tools')
            .addItem('Show Purchasing Sidebar', 'purchasingSidebar')
            .addSeparator()
            .addSubMenu(statusUpdateSubMenu)
            .addSeparator()
            .addItem('Clean Up Formats, Formulae and Drop-downs', 'matCleanup')
            .addToUi();
        setupCLPicker();
        setupGoodsReceiver();
        matDropDowns();
    }
    function purchasingSidebar() {
        UI = SpreadsheetApp.getUi();
        var html = HtmlService.createHtmlOutputFromFile('SideBarPlain')
            .setSandboxMode(HtmlService.SandboxMode.IFRAME)
            .setTitle('Purchasing Side Bar')
            .setWidth(250);
        UI.showSidebar(html);
    }
    function onInstall(e) {
        onOpen(e);
    }
    function edited(e) {
        Logger.log("Cell Edited = " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getA1Notation());
        UI = SpreadsheetApp.getUi();
        SH = SS.getActiveSheet();
        var SHname = SH.getName();
        thisCell = SH.getActiveCell();
        if (SHname == CORELIST_SHEET) {
            clEdit(e);
        }
        if (SHname == MATERIALS_SHEET) {
            matEdit(e);
        }
        if (SHname == GOODS_RECEIVING_SHEET) {
            grEdit(e);
        }
        if (SHname == SVS_MATCHER_SHEET) {
            svsEdit(e);
        }
        if (SHname == NONCORE_SHEET) {
            ncEdit(e);
        }
    }
})(Code || (Code = {}));
