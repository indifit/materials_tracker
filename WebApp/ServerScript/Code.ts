var hash: string;

function doGet(request: GoogleAppsScript.Script.IParameters)
{
    var pageSelector: jw.MaterialsTracker.Utilities.PageSelector = new jw.MaterialsTracker.Utilities.PageSelector(request);

    var page: jw.MaterialsTracker.Interfaces.IPage = pageSelector.getPage();    

    var template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(page.templateName);

    template.data = page.data;

    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getCoreListData(filter: jw.MaterialsTracker.Interfaces.ICoreListFilter) : jw.MaterialsTracker.Interfaces.ICoreListData
{
    var dataFetcher: jw.MaterialsTracker.Utilities.DataFetcher = new jw.MaterialsTracker.Utilities.DataFetcher();

    var filteredCoreListObjects: Object[] = dataFetcher.getFilteredCoreListItems(filter);

    var categories: string[] = [];

    var types: string[] = [];       

    //If no filter has been passed only retrieve the trades
    if (typeof filter != 'undefined')
    {        
        //Retrieve the sub categories associated with the selected trade
        filteredCoreListObjects.forEach((value: any): void =>
        {
            if (categories.indexOf(value.productSubCategory.toString().trim()) === -1)
            {
                categories.push(value.productSubCategory.toString().trim());
            }
        });
        

        if (typeof filter.category != 'undefined')
        {
            //Retrieve the types for the category selected
            filteredCoreListObjects.forEach((value: any): void =>
            {
                if (types.indexOf(value.type.toString().trim()) === -1)
                {
                    types.push(value.type.toString().trim());
                }
            });
        }

        return {
            filteredListData: filteredCoreListObjects,
            subCategories: categories,
            types: types
        };

    } else
    {
        return {};
    }
       
}