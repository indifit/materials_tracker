var hash: string;

function doGet(request: GoogleAppsScript.Script.IParameters)
{
    var pageSelector: jw.MaterialsTracker.Utilities.PageSelector = new jw.MaterialsTracker.Utilities.PageSelector(request);

    var page: jw.MaterialsTracker.Interfaces.IPage = pageSelector.getPage();    

    var template = HtmlService.createTemplateFromFile(page.templateName);

    template.data = page.data;

    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getSavedItems(projectSsid: string): jw.MaterialsTracker.Interfaces.ISavedItem[]
{
    var projectSs: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(projectSsid);       

    var materialsTrackingSheet: GoogleAppsScript.Spreadsheet.Sheet = projectSs.getSheetByName('Materials Tracking');

    var savedItemsRange: GoogleAppsScript.Spreadsheet.Range = materialsTrackingSheet.getRange(2, 1, materialsTrackingSheet.getLastRow(), materialsTrackingSheet.getLastColumn());

    var savedItemsValues = savedItemsRange.getValues();

    var rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(savedItemsValues);

    var savedItems: any[] = rangeUtils.convertToObjectArray();

    var ret: jw.MaterialsTracker.Interfaces.ISavedItem[] = [];

    for (var i = 0; i < savedItems.length; i++)
    {
        if (savedItems[i].itemCode !== '')
        {
            ret.push({
                itemCode: savedItems[i].itemCode.toString(),
                pdc: savedItems[i].pdCode.toString(),
                quantity: parseInt(savedItems[i].qty.toString())
            });   
        }        
    }

    return ret;
}

function getCoreListData(filter: jw.MaterialsTracker.Interfaces.ICoreListFilter) : jw.MaterialsTracker.Interfaces.ICoreListData
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

    var filteredListData: Object[] = rangeUtils.convertToObjectArray();

    filteredListData.forEach((value: any, index: number, arr: Object[]): void => {
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

        filteredListData = rangeUtils.convertToObjectArray();
        
        //Retrieve the sub categories associated with the selected trade
        filteredListData.forEach((value: any): void =>
        {
            if (subCategories.indexOf(value.productSubCategory.toString().trim()) === -1)
            {
                subCategories.push(value.productSubCategory.toString().trim());
            }
        });

        if (typeof filter.category != 'undefined')
        {
            //Filter the rows again to those for the selected category
            filteredRows = jw.MaterialsTracker.Utilities.RangeUtilties.findRowsMatchingKey(filteredRows, filter.category, 3, headerRow);

            rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(filteredRows);

            filteredListData = rangeUtils.convertToObjectArray();

            //Retrieve the types for the category selected
            filteredListData.forEach((value: any): void =>
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

            filteredListData = rangeUtils.convertToObjectArray();                        
        }

        //Retrieve the valid project dimensions for the filtered items
        var allpdcsSheet = centralPurchasingSS.getSheetByName('PDCs');

        var allpdcsRange = allpdcsSheet.getRange(2, 1, allpdcsSheet.getLastRow(), allpdcsSheet.getLastColumn());        

        filteredListData.forEach((value: any): void =>
        {
            var pdcString = value.pdc;

            //Get the pdc codes for this item
            var pdcArray = pdcString.split(';');

            value.pdcs = [];

            pdcArray.forEach((pdcCode: any): void =>
            {
                var row: Object[] = jw.MaterialsTracker.Utilities.RangeUtilties.findFirstRowMatchingKey(allpdcsRange, pdcCode);

                if (row != null) {
                    value.pdcs.push({ code: pdcCode, description: row[1].toString() });
                }                                
            });
        });

        return {
            coreListData: coreListData,
            filteredListData: filteredListData,
            trades: trades,
            subCategories: subCategories,
            types: types
        };

    } else
    {
        return {    
            coreListData: coreListData,        
            trades: trades
        };
    }
       
}