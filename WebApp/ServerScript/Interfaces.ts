module jw.MaterialsTracker.Interfaces {
    /*
     * Interface that represents a mapping of a randomly generated
     * ID to a valid project number
     * This type will be returned from a function running on the
     * server that will return all such mappings
     */
    export interface IProjectHashLookupResponse {
        projectNumber: number;
        projectName: string;
        kingdomHallAddress?: string;
        urlHash: string;
    }

    export interface IPage
    {
        templateName: string;
        data: Object;
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