var Materials;
(function (Materials) {
    var matSH = SS.getSheetByName(MATERIALS_SHEET);
    var matFirst = MHRI + 2;
    var matLast = matSH.getLastRow();
    function matEdit(e) {
        if (e.range.getColumn() == col(H_SUPPLIER).n) {
            e.range.clearNote();
            matSH.getRange(col(H_ADMINCODE).a + e.range.getRow()).setValue('');
            var supplier = getSupplier(e.value);
            if (supplier.name == e.value) {
                var hoverNote = supplier.contact + '\n' + supplier.tel + '\n' + supplier.email;
                e.range.setNote(hoverNote);
                matSH.getRange(col(H_ADMINCODE).a + e.range.getRow()).setValue(supplier.admin);
            }
        }
        if (e.range.getColumn() == col(H_STATUS).n) {
            var thisPDN = matSH.getRange(col(H_PDN).a + e.range.getRow());
            var thisPDC = matSH.getRange(col(H_PDC).a + e.range.getRow());
            if (e.value.substr(0, 1) >= VOID_PREFIX) {
                if (thisPDN.getNote() == '') {
                    thisPDN.setNote('was: ' + thisPDN.getValue());
                    thisPDC.setNote('was: ' + thisPDC.getValue());
                    thisPDN.setValue('void');
                }
            }
            if (e.value.substr(0, 1) < VOID_PREFIX) {
                if (thisPDN.getNote() != '') {
                    thisPDN.setValue(thisPDN.getNote().substr(5));
                    thisPDN.clearNote();
                    thisPDC.clearNote();
                }
            }
        }
    }
    function matCleanup() {
        matDropDowns();
        matDateFormat();
        matNumberFormat();
        matSetFormulae();
    }
    function matDropDowns() {
        var CDD = getCentralDropDowns();
        var allSuppliers = getSupplier('*all*');
        var supplierNames = [];
        for (; allSuppliers.length > 0;) {
            supplierNames.unshift(allSuppliers.pop()[0]);
        }
        supplierNames.shift();
        var pdnDDSource = CDD.ddPDN;
        var columns = [
            H_TYPE, CDD.ddType,
            H_TEAM, CDD.ddTeams,
            H_PUOM, CDD.ddUoM,
            H_BUOM, CDD.ddUoM,
            H_PDN, pdnDDSource,
            H_VATRATE, CDD.ddVATRates,
            H_SUPPLIER, supplierNames,
            H_EMERGENCY, ['Yes', 'No'],
            H_STATUS, CDD.ddStatusBPR,
            H_HIRESTATUS, CDD.ddStatusHire
        ];
        for (; columns.length > 0;) {
            var thisCol = columns.shift();
            var colRange = col(thisCol).a + (MHRI + 2) + ':' + col(thisCol).a;
            var dv = SpreadsheetApp.newDataValidation();
            dv.setAllowInvalid(false);
            var _dd = columns.shift();
            dv.requireValueInList(_dd, true);
            SS.getSheetByName(MATERIALS_SHEET).getRange(colRange).setDataValidation(dv.build());
        }
    }
    function matDateFormat() {
        var dtColumns = [
            col(H_ACTDEL).a,
            col(H_OFF).a,
            col(H_POCREATED).a
        ];
        for (; dtColumns.length > 0;) {
            var _col = dtColumns.pop();
            var fRange = matSH.getRange(_col + matFirst + ":" + _col + matLast);
            var fDT = [];
            var fD = [];
            for (var i = fRange.getNumRows(); i > 0; i--) {
                fDT.push(["dd/MM/yyyy HH:mm:ss"]);
                fD.push(["dd/MM/yyyy"]);
            }
            fRange.setNumberFormats(fDT);
            fRange.setNumberFormats(fD);
            fRange.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
        }
    }
    function matNumberFormat() {
        var colNumFormat = [
            col(H_ICODE).a, '@STRING@',
            col(H_UNIT).a, '0.00',
            col(H_NET).a, '0.00',
            col(H_VATRATE).a, '0%',
            col(H_VATVALUE).a, '0.00',
            col(H_LINEVALUE).a, '0.00'
        ];
        for (; colNumFormat.length > 0;) {
            var _col = colNumFormat.shift();
            var _format = colNumFormat.shift();
            var fRange = matSH.getRange(_col + matFirst + ":" + _col + matLast);
            var fNum = [];
            for (var i = fRange.getNumRows(); i > 0; i--) {
                fNum.push([_format]);
            }
            fRange.setNumberFormats(fNum);
        }
    }
    function matSetFormulae() {
        var formulaColumns = [
            col(H_NET).a,
            '=IF(AND(EQ(' + col(H_QTY).a + matFirst + ',""),EQ(' + col(H_UNIT).a + matFirst + ',"")),"",' + col(H_QTY).a + matFirst + '*' + col(H_UNIT).a + matFirst + ')',
            col(H_VATVALUE).a,
            '=IF(EQ(' + col(H_VATRATE).a + matFirst + ',""),"",' + col(H_NET).a + matFirst + '*(RIGHT(' + col(H_VATRATE).a + matFirst + ',(LEN(' + col(H_VATRATE).a + matFirst + ')-2))))',
            col(H_LINEVALUE).a,
            '=IF(AND(EQ(' + col(H_NET).a + matFirst + ',"")),"",' + col(H_NET).a + matFirst + '+' + col(H_VATVALUE).a + matFirst + ')',
            col(H_PDC).a,
            '=IF(EQ(' + col(H_PDN).a + matFirst + ',""),"",VLOOKUP(' + col(H_PDN).a + matFirst + ',PD_Range,2,FALSE))',
            col(H_QTYLEFT).a,
            '=IF(EQ(' + col(H_QTY).a + matFirst + ',""),"",' + col(H_QTY).a + matFirst + '-' + col(H_QTYRCVD).a + matFirst + ')'
        ];
        for (; formulaColumns.length > 0;) {
            var initCol = formulaColumns.shift();
            var initRange = matSH.getRange(initCol + matFirst);
            initRange.setFormula(formulaColumns.shift());
            initRange.copyTo(matSH.getRange(initCol + (matFirst + 1) + ":" + initCol + matLast));
        }
    }
    var hRow = null;
    function col(columnHeader) {
        if (!hRow) {
            hRow = matSH.getRange((MHRI + 1), 1, 1, matSH.getLastColumn()).getValues();
        }
        var hIndex = hRow[0].indexOf(columnHeader);
        if (hIndex > -1) {
            var retVal = { i: hIndex, n: (hIndex + 1), a: alphaCols[hIndex] };
            return retVal;
        }
    }
    function updateStatustoSent() {
        UI = SpreadsheetApp.getUi();
        var thisRowIndex = SH.getActiveCell().getRow() - matFirst;
        if (thisRowIndex <= MHRI) {
            var wrongRow = UI.alert("Status Update Error", "You have currently selected one of the header rows."
                + "\nPlease select a cell in a row from the materials list.", UI.ButtonSet.OK);
            return false;
        }
        if (SH.getName() != MATERIALS_SHEET) {
            var uiResponse = UI.alert("Status Update Error", "You are not on the " + MATERIALS_SHEET + " sheet.", UI.ButtonSet.OK);
            return false;
        }
        var statusBPRReady = '3 BPR Prepared';
        var statusBPRSent = '4 BPR Sent';
        var mtPRNoRange = matSH.getRange(matFirst, col(H_PONUM).n, matLast, 1);
        var mtPRNoVals = mtPRNoRange.getValues();
        var mtStatusRange = matSH.getRange(matFirst, col(H_STATUS).n, matLast, 1);
        var mtStatusVals = mtStatusRange.getValues();
        var thisBPRNo = mtPRNoVals[thisRowIndex][0];
        var okToUpdate = true;
        for (var u = 0; u < mtPRNoVals.length; u++) {
            if (mtPRNoVals[u][0] == thisBPRNo) {
                if (mtStatusVals[u][0] != statusBPRReady) {
                    okToUpdate = false;
                }
                if (mtStatusVals[u][0] == statusBPRReady) {
                    mtStatusVals[u][0] = statusBPRSent;
                }
            }
            if (!okToUpdate) {
                break;
            }
        }
        if (!okToUpdate) {
            var uiResponse = UI.alert("Status Update Error", "Not all line items for BPR " + thisBPRNo + " have a status of \"" + statusBPRReady + "\".", UI.ButtonSet.OK);
            return false;
        }
        if (okToUpdate) {
            mtStatusRange.setValues(mtStatusVals);
        }
    }
})(Materials || (Materials = {}));
