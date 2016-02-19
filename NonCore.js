var ncHeaderRow = 2;
var ddProcessNonCore = 'G2';
var NonCore;
(function (NonCore) {
    function ncEdit(e) {
        var eRangeA1 = e.range.getA1Notation();
        if (eRangeA1 == ddProcessNonCore && e.value == 'GO!') {
            SS.toast("Please wait...", "Working", 600);
            processNonCore();
            e.range.setValue('Add to Materials Tracker');
        }
    }
    function processNonCore() {
        SS.toast("Processing Non Core Items", "Non Core 1 of 3", 600);
        var ncSH = SS.getSheetByName(NONCORE_SHEET);
        var ncRange = ncSH.getRange((ncHeaderRow + 1), 1, ncSH.getLastRow(), ncSH.getLastColumn());
        var ncValues = ncRange.getValues();
        var ncToAdd = [];
        for (var n = 0; n < ncValues.length; n++) {
            var _ncToAdd = [];
            var _ncIC = '';
            var _ncID = '';
            var _ncType = COREEXTRA;
            var _ncTemp = '';
            if (ncValues[n][0] != '') {
                _ncTemp = ncValues[n][0].toString().trim();
                _ncIC = '(' + _ncTemp + ')';
                _ncID = _ncIC;
            }
            if (ncValues[n][1] != '') {
                _ncTemp = ncValues[n][1].toString().trim().toUpperCase();
                if (_ncID != '') {
                    _ncID += ' ';
                }
                _ncID += _ncTemp;
            }
            if (ncValues[n][2] != '') {
                _ncTemp = ncValues[n][2].toString().trim().toUpperCase();
                if (_ncID != '') {
                    _ncID += ' ';
                }
                _ncID += _ncTemp;
            }
            if (ncValues[n][3] != '') {
                _ncTemp = ncValues[n][3].toString().trim().toUpperCase();
                if (_ncID != '') {
                    _ncID += ' ';
                }
                _ncID += _ncTemp;
            }
            if (ncValues[n][4] != '') {
                _ncTemp = ncValues[n][4].toString().trim().toUpperCase();
                if (_ncID != '') {
                    _ncID += ' ';
                }
                _ncID += _ncTemp + ':';
            }
            if (ncValues[n][5] != '') {
                _ncTemp = ncValues[n][5].toString().trim().toLowerCase();
                if (_ncID != '') {
                    _ncID += ' ';
                }
                _ncID += _ncTemp;
            }
            if (_ncID != '') {
                _ncToAdd.push(_ncID);
                _ncToAdd.push('');
                _ncToAdd.push(_ncType);
                ncToAdd.push(_ncToAdd);
            }
        }
        SS.toast("Adding Non Core Items", "Non Core 2 of 3", 600);
        matSH = SS.getSheetByName(MATERIALS_SHEET);
        var curMatRange = matSH.getRange(1, 2, matSH.getLastRow(), matSH.getLastColumn());
        var curMat = curMatRange.getValues();
        var blankRow = 0;
        for (var i = curMat.length - 1; i > MHRI; i--) {
            if ((curMat[i][curMat[MHRI].indexOf(H_IDESC)] != '')
                || (curMat[i][curMat[MHRI].indexOf(H_ICODE)] != '')
                || (curMat[i][curMat[MHRI].indexOf(H_TYPE)] != '')) {
                blankRow = i + 2;
                break;
            }
        }
        matSH.getRange(blankRow, col(H_IDESC).n, ncToAdd.length, ncToAdd[0].length).setValues(ncToAdd);
        SS.toast("Clearing Non Core (Free Text)", "Non Core 3 of 3", 5);
        ncRange.clear();
    }
})(NonCore || (NonCore = {}));
