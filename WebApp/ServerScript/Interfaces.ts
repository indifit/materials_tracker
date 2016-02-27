module jw.MaterialsTracker.Interfaces {
    /*
     * Interface that represents a mapping of a randomly generated
     * ID to a valid project number
     * This type will be returned from a function running on the
     * server that will return all such mappings
     */
    export interface IProjectHashLookupResponse {
        projectNumber: number;
        urlHash: string;
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