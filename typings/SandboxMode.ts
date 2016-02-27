// ReSharper disable once InconsistentNaming
module GoogleAppsScript.HTML
{
    export interface ISandboxModeType
    {                
    }

    export class SandboxMode
    {        
        static EMULATED: ISandboxModeType; 
        static IFRAME: ISandboxModeType;
        static NATIVE: ISandboxModeType;
    }
}