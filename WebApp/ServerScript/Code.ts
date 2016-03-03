var hash: string;

function doGet(request: GoogleAppsScript.Script.IParameters)
{
    var pageSelector: jw.MaterialsTracker.Utilities.PageSelector = new jw.MaterialsTracker.Utilities.PageSelector(request);

    var page: jw.MaterialsTracker.Interfaces.IPage = pageSelector.getPage();    

    var template = HtmlService.createTemplateFromFile(page.templateName);

    template.data = page.data;

    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getCoreListData(filter: jw.MaterialsTracker.Interfaces.ICoreListFilter) : any
{
    var centralPurchasingSSID: string = jw.MaterialsTracker.Config.ConfigurationManager.getSetting('CentralPurchasingSSID');

    var centralPurchasingSS = SpreadsheetApp.openById(centralPurchasingSSID);

    var coreListSheet = centralPurchasingSS.getSheetByName('CoreList');

    var lastRow: number = coreListSheet.getLastRow();

    var lastColumn: number = coreListSheet.getLastColumn();

    var coreListRange: GoogleAppsScript.Spreadsheet.Range = coreListSheet.getRange(1, 1, lastRow, lastColumn);

    var filteredRows: Object[][] = null;

    var processedTrades: string[] = [];

    //Get the trades from the core list
    var rangeUtils: jw.MaterialsTracker.Utilities.RangeUtilties;

    rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(coreListRange);

    var coreListData: Object[] = rangeUtils.convertToObjectArray();

    coreListData.forEach((value: any, index: number, arr: Object[]): void => {
        if (processedTrades.indexOf(value.trade.toString().trim()) === -1) {
            processedTrades.push(value.trade.toString().trim());
        }
    });

    //If no filter has been passed only retrieve the trades
    if (typeof filter != 'undefined')
    {
        filteredRows = jw.MaterialsTracker.Utilities.RangeUtilties.findRowsMatchingKey(coreListRange, filter.trade, 0, true);
    } else
    {
        return {            
            trades: processedTrades
        };
    }
    
    if (filteredRows == null)
    {
        rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(coreListRange);
    } else
    {
        rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(filteredRows);
    }

    coreListData = rangeUtils.convertToObjectArray();                   

    return {
        coreListData: coreListData,
        trades: processedTrades
    };
}