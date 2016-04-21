var hash;
function doGet(request) {
    var pageSelector = new jw.MaterialsTracker.Utilities.PageSelector(request);
    var page = pageSelector.getPage();
    var template = HtmlService.createTemplateFromFile(page.templateName);
    template.data = page.data;
    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
function getCoreListData(filter) {
    var dataFetcher = new jw.MaterialsTracker.Utilities.DataFetcher();
    var filteredCoreListObjects = dataFetcher.getFilteredCoreListItems(filter);
    var categories = [];
    var types = [];
    if (typeof filter != 'undefined') {
        filteredCoreListObjects.forEach(function (value) {
            if (categories.indexOf(value.productSubCategory.toString().trim()) === -1) {
                categories.push(value.productSubCategory.toString().trim());
            }
        });
        if (typeof filter.category != 'undefined') {
            filteredCoreListObjects.forEach(function (value) {
                if (types.indexOf(value.type.toString().trim()) === -1) {
                    types.push(value.type.toString().trim());
                }
            });
        }
        return {
            filteredListData: filteredCoreListObjects,
            subCategories: categories,
            types: types
        };
    }
    else {
        return {};
    }
}
function saveBasketToMaterialsTracker(basketItems, projectDetails) {
    jw.MaterialsTracker.Utilities.DataSaver.saveBasketData(projectDetails, basketItems);
}
