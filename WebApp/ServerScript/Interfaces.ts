module jw.MaterialsTracker.Interfaces {
    /*
     * Interface that represents a mapping of a randomly generated
     * ID to a valid project number
     * This type will be returned from a function running on the
     * server that will return all such mappings
     */
    export interface IProjectHashLookupResponse {
        projectNumber: string;
        projectName: string;
        urlHash: string;
        projectSsid?: string;
    }

    export interface IPage
    {
        templateName: string;
        data: Object;
    }

    export interface ICoreListFilter
    {
        trade?: string;
        category?: string;
        type?: string;
    }

    export interface ICoreListData
    {        
        listData?: any[];
        subCategories?: string[];
        types?: string[];
    }    

    export interface ISavedItem
    {
        itemCode: string;
        quantity: number;
        pdc: string;
    }
} 

module GoogleAppsScript.Script
{
    export interface IParameters
    {
        queryString: string;
        parameter: { [key: string]: string };
        parameters: {[key: string]: string[]};
    }
}