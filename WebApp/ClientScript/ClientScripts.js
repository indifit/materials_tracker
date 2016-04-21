var jw;
(function (jw) {
    var MaterialsTracker;
    (function (MaterialsTracker) {
        var Client;
        (function (Client) {
            var CoreListData = (function () {
                function CoreListData(data) {
                    this.listData = data;
                }
                return CoreListData;
            }());
            Client.CoreListData = CoreListData;
            var Filter = (function () {
                function Filter(trade, category, type) {
                    if (typeof trade != 'undefined') {
                        this.trade = trade;
                    }
                    if (typeof category != 'undefined') {
                        this.category = category;
                    }
                    if (typeof type != 'undefined') {
                        this.type = type;
                    }
                }
                return Filter;
            }());
            Client.Filter = Filter;
            var CoreListFilter = (function () {
                function CoreListFilter(coreListData) {
                    var _this = this;
                    this.filterCoreList = function (filter) {
                        var filteredData = {
                            listData: _this.coreListData.listData,
                            types: [],
                            subCategories: []
                        };
                        if (typeof filter != 'undefined' && filter != null) {
                            if (typeof filter.trade != 'undefined') {
                                filteredData.listData = [];
                                for (var i = 0; i < _this.coreListData.listData.length; i++) {
                                    if (_this.coreListData.listData[i].trade.toString().trim().toLowerCase() === filter.trade.trim().toLowerCase()) {
                                        filteredData.listData.push(_this.coreListData.listData[i]);
                                        if (filteredData.subCategories.indexOf(_this.coreListData.listData[i].productSubCategory) === -1) {
                                            filteredData.subCategories.push(_this.coreListData.listData[i].productSubCategory);
                                        }
                                    }
                                }
                            }
                            if (typeof filter.category != 'undefined' && filter.category != null) {
                                var tempListData = [];
                                for (var i = 0; i < filteredData.listData.length; i++) {
                                    if (filteredData.listData[i].productSubCategory.toString().trim().toLowerCase() === filter.category.trim().toLowerCase()) {
                                        tempListData.push(filteredData.listData[i]);
                                    }
                                }
                                filteredData.listData = tempListData;
                                for (var i = 0; i < filteredData.listData.length; i++) {
                                    if (filteredData.types.indexOf(filteredData.listData[i].type) === -1) {
                                        filteredData.types.push(filteredData.listData[i].type);
                                    }
                                }
                            }
                            if (typeof filter.type != 'undefined' && filter.type !== null) {
                                tempListData = [];
                                for (var i = 0; i < filteredData.listData.length; i++) {
                                    if (filteredData.listData[i].type.toString().trim().toLowerCase() === filter.type.trim().toLowerCase()) {
                                        tempListData.push(filteredData.listData[i]);
                                    }
                                }
                                filteredData.listData = tempListData;
                            }
                        }
                        return filteredData;
                    };
                    this.coreListData = coreListData;
                }
                return CoreListFilter;
            }());
            Client.CoreListFilter = CoreListFilter;
        })(Client = MaterialsTracker.Client || (MaterialsTracker.Client = {}));
    })(MaterialsTracker = jw.MaterialsTracker || (jw.MaterialsTracker = {}));
})(jw || (jw = {}));
ko.utils.extendObservable = function (target, source) {
    var prop, srcVal, isObservable = false;
    for (prop in source) {
        if (!source.hasOwnProperty(prop)) {
            continue;
        }
        if (ko.isWriteableObservable(source[prop])) {
            isObservable = true;
            srcVal = source[prop]();
        }
        else if (typeof (source[prop]) !== 'function') {
            srcVal = source[prop];
        }
        if (ko.isWriteableObservable(target[prop])) {
            target[prop](srcVal);
        }
        else if (target[prop] === null || target[prop] === undefined) {
            target[prop] = isObservable ? ko.observable(srcVal) : srcVal;
        }
        else if (typeof (target[prop]) !== 'function') {
            target[prop] = srcVal;
        }
        isObservable = false;
    }
};
ko.utils.clone = function (obj, emptyObj) {
    var json = ko.toJSON(obj);
    var js = JSON.parse(json);
    ko.utils.extendObservable(emptyObj, js);
};
