/* GENERAL */
var SS = SpreadsheetApp.getActiveSpreadsheet();
var SH = SS.getActiveSheet();
var thisCell;
var UI = null;

/* Drive Owner User ID */
function DRIVE_OWNER() {
    return getCentralDropDowns().DRIVE_OWNER[0][0];
}

var WT_PREFIX = 'WT50';

function PROJ_NUMBER() {
    return SS.getRangeByName('projNumber').getValue().toString();
}
function PROJ_NAME() {
    return SS.getRangeByName('projName').getValue().toString();
}

var CORELIST_SHEET = 'Core List';
var NONCORE_SHEET = 'Non Core (Free Text)';
var MATERIALS_SHEET = 'Materials Tracking';
var SVS_MATCHER_SHEET = 'SVS Matcher';
var GOODS_RECEIVING_SHEET = 'Goods Receiving';
var LOCAL_SUPPLIERS_SHEET = 'Project-specific Suppliers';

/* MHRI = Materials Header Row Index */
var MHRI = 1;

function CentralData() {
    var retVal = SS.getRangeByName('settings_CD').getValue().toString();
    return retVal;
}

function PD(PDNorPDC) {
    var dds = getCentralDropDowns();
    var PDNs = dds.ddPDN;
    var PDCs = dds.ddPDC;

    var retVal = PDCs[PDNs.indexOf(PDNorPDC)];
    if (!retVal) {
        retVal = PDNs[PDCs.indexOf(PDNorPDC)];
    }
    if (!retVal) {
        retVal = null;
    }

    return retVal;
}

var CD_WTSuppliers = 'WTSuppliers';
var CD_CoreList = 'CoreList';

/* Quotes */
function QUOTE_TEMPLATE_ID() {
    return getCentralDropDowns().TEMPLATE_IDS[2][0];
}

/* PO 3 */
function PO_TEMPLATE_ID() {
    return getCentralDropDowns().TEMPLATE_IDS[0][0];
}

function PO_FOLDER_ID() {
    return SS.getRangeByName('settings_POFolder').getValue().toString();
}

/* PR 3 */
function PR_TEMPLATE_ID() {
    return getCentralDropDowns().TEMPLATE_IDS[1][0];
}
function PR_FOLDER_ID() {
    return SS.getRangeByName('settings_PRFolder').getValue().toString();
}

/* SVS */
function SVS_FOLDER_ID() {
    var retVal = {
        toDo: SS.getRangeByName('settings_SVStoDoFolder').getValue().toString()
    };
    return retVal;
}

/* PR Emails */
var PRemailTo = "sep@talktalk.net";
var PRemailCC = "dean@longs-home.co.uk";

/* Material Types */
var MATERIALS = 'WT50 Material';
var HIRE = 'Core Hire';
var CORELIST = 'Core Item';
var COREEXTRA = 'Non Core';
var TYPE = '';

/* POPR 0 */
var POorPR = 'PO';
var OrderOrRequest = 'Order';

var PR_SUPPLIER = 'Branch Purchasing';
var PR_ITEMS_LIMIT = 20;

/* Useful stuff */
var alphaCols = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
    'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X',
    'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH',
    'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ',
    'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];

// a and b are javascript Date objects
function dateDiffInDays(a, b) {
    // Discard the time and time-zone information.
    var utcA = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
    var utcB = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());

    var _MS_PER_DAY = (1000 * 60 * 60 * 24);

    return Math.floor((utcA - utcB) / _MS_PER_DAY);
}

/* MATERIALS COLUMN HEADERS */
// Materials initial columns
var H_LINEID = 'Line ID';

// used on PO / PR
var H_IDESC = 'Item Description';
var H_ICODE = 'Item Code';

var H_TYPE = 'Type';

// used for PR
var H_NOTES = 'Notes or URL for Core extras';

var H_USAGE = 'Usage';
var H_TEAM = 'Team';

// used on PO / PR
var H_QTY = 'Qty';
var H_PUOM = 'Purchase UoM';
var H_FACTOR = 'Factor';
var H_BUOM = 'Base UoM';
var H_UNIT = 'Unit Cost';

var H_NET = 'Net Cost';
var H_VATRATE = 'VAT Rate';
var H_VATVALUE = 'VAT';
var H_LINEVALUE = 'Total';

var H_PDN = 'Project Dimension';
var H_PDC = 'PD Code';

// order details
var H_SUPPLIER = 'Supplier';
var H_EMERGENCY = 'Emergency Order';
var H_ACTDEL = 'Delivery to Site Date';
var H_OFF = 'Off Hire Date (if Hire)';
var H_PONUM = 'Auto generated PR Or PO Number';
var H_POCREATED = 'Auto Date Order Created';
var H_BRANCHPO = 'Branch PO Number\n(link to SVS)';
var H_BRANCHLINE = 'Branch PO Line number';
var H_BRANCHSUPPLIER = 'Branch PO Supplier';

var H_STATUS = 'Order Status';

var ITEMREADY_PREFIX = 2;
var REGEN_PREFIX = 3;
var EXPECTINGSVS_PREFIX = 4;
var VOID_PREFIX = 8;

function SVS_RECEIVED() {
    return getCentralDropDowns().ddStatusBPR[4].toString();
}

// delivery details
var H_QTYRCVD = 'Qty Received';
var H_QTYLEFT = 'Sum Left to deliver';
var H_DATEDELIVERED = 'Date Delivered';
var H_STORED = 'Store location';
var H_DELCOMMENTS = 'Delivery queries or notes about paying';

// hire items
var H_HIRESTATUS = 'Hire Status';
var H_ONHIRE = 'Date delivered to site';
var H_OFFHIRE = 'Off hire date';
var H_HIRECOLLECTION = 'Collection from site date';
//# sourceMappingURL=Statics and Vars.js.map
