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

    var trades: string[] = [];

    var subCategories: string[] = [];

    var types: string[] = [];

    var rangeValues: Object[][] = coreListRange.getValues();

    var headerRow: Object[] = rangeValues[0];

    //Get the trades from the core list
    var rangeUtils: jw.MaterialsTracker.Utilities.RangeUtilties;

    rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(coreListRange);

    var coreListData: Object[] = rangeUtils.convertToObjectArray();

    coreListData.forEach((value: any, index: number, arr: Object[]): void => {
        if (trades.indexOf(value.trade.toString().trim()) === -1) {
            trades.push(value.trade.toString().trim());
        }
    });

    //If no filter has been passed only retrieve the trades
    if (typeof filter != 'undefined')
    {
        //Filter the rows to those for the selected trade
        filteredRows = jw.MaterialsTracker.Utilities.RangeUtilties.findRowsMatchingKey(rangeValues, filter.trade, 0, headerRow);

        rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(filteredRows);

        coreListData = rangeUtils.convertToObjectArray();
        
        //Retrieve the sub categories associated with the selected trade
        coreListData.forEach((value: any): void =>
        {
            if (subCategories.indexOf(value.productsubcategory.toString().trim()) === -1)
            {
                subCategories.push(value.productsubcategory.toString().trim());
            }
        });

        if (typeof filter.category != 'undefined')
        {
            //Filter the rows again to those for the selected category
            filteredRows = jw.MaterialsTracker.Utilities.RangeUtilties.findRowsMatchingKey(filteredRows, filter.category, 3, headerRow);

            rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(filteredRows);

            coreListData = rangeUtils.convertToObjectArray();

            //Retrieve the types for the category selected
            coreListData.forEach((value: any): void =>
            {
                if (types.indexOf(value.type.toString().trim()) === -1)
                {
                    types.push(value.type.toString().trim());
                }
            });
        }

        if (typeof filter.type != 'undefined') {
            //Filter the rows again to those for the selected type
            filteredRows = jw.MaterialsTracker.Utilities.RangeUtilties.findRowsMatchingKey(filteredRows, filter.type, 4, headerRow);

            rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(filteredRows);

            coreListData = rangeUtils.convertToObjectArray();            
        }

        return {
            coreListData: coreListData,
            trades: trades,
            subCategories: subCategories,
            types: types
        };

    } else
    {
        return {            
            trades: trades
        };
    }
       
}