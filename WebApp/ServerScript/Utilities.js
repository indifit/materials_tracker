var jw;
(function (jw) {
    (function (MaterialsTracker) {
        (function (Utilities) {
            var RangeUtilties = (function () {
                function RangeUtilties() {
                }
                RangeUtilties.findFirstRowMatchingKey = function (range, lookupVal, keyColIndex) {
                    if (typeof keyColIndex === "undefined") { keyColIndex = 0; }
                    var vals = range.getValues();
                    var rowVals = null;
                    for (var i = 0; i < vals.length; i++) {
                        rowVals = vals[i];
                        var keyColVal = rowVals[keyColIndex];

                        Logger.log('keyColVal = ' + rowVals[keyColIndex]);

                        if (typeof keyColVal != "undefined" && keyColVal.toString().toLowerCase() === lookupVal.toLowerCase()) {
                            return rowVals;
                        }
                    }
                    return rowVals;
                };
                return RangeUtilties;
            })();
            Utilities.RangeUtilties = RangeUtilties;
        })(MaterialsTracker.Utilities || (MaterialsTracker.Utilities = {}));
        var Utilities = MaterialsTracker.Utilities;
    })(jw.MaterialsTracker || (jw.MaterialsTracker = {}));
    var MaterialsTracker = jw.MaterialsTracker;
})(jw || (jw = {}));
//# sourceMappingURL=Utilities.js.map
