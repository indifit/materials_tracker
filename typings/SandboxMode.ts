module GoogleAppsScript.HTML
{
    export interface ISandboxModeType
    {                
    }

    export class SandboxMode
    {        
        EMULATED: ISandboxModeType; 
        IFRAME: ISandboxModeType;
        NATIVE: ISandboxModeType;
    }
}