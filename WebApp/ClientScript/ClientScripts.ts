module jw.MaterialsTracker.Client
{
    export class CoreListData implements MaterialsTracker.Interfaces.ICoreListData
    {
        constructor(data: any[])
        {
            this.listData = data;
        }

        listData: any[];
        subCategories: string[];
        types: string[];
    }

    export class Filter implements MaterialsTracker.Interfaces.ICoreListFilter
    {
        constructor(trade?: string, category?: string, type?: string)
        {
            if (typeof trade != 'undefined')
            {
                this.trade = trade;    
            }
            
            if (typeof category != 'undefined')
            {
                this.category = category;
            }

            if (typeof type != 'undefined')
            {
                this.type = type;
            }
        }

        trade: string;
        category: string;
        type: string;
    }

    export class CoreListFilter
    {
        private coreListData: MaterialsTracker.Interfaces.ICoreListData;

        constructor(coreListData: MaterialsTracker.Interfaces.ICoreListData)
        {
            this.coreListData = coreListData;
        }

        public filterCoreList = (filter: MaterialsTracker.Interfaces.ICoreListFilter) :  MaterialsTracker.Interfaces.ICoreListData =>
        {
            var filteredData: MaterialsTracker.Interfaces.ICoreListData = {
                listData: this.coreListData.listData,
                types: [],
                subCategories: []
            };

            if (typeof filter != 'undefined' && filter != null)
            {
                if (typeof filter.trade != 'undefined')
                {
                    filteredData.listData = [];

                    for (var i = 0; i < this.coreListData.listData.length; i++)
                    {
                        if (this.coreListData.listData[i].trade.toString().trim().toLowerCase() === filter.trade.trim().toLowerCase())
                        {
                            filteredData.listData.push(this.coreListData.listData[i]);

                            if (filteredData.subCategories.indexOf(this.coreListData.listData[i].productSubCategory) === -1)
                            {
                                filteredData.subCategories.push(this.coreListData.listData[i].productSubCategory);
                            }
                        }
                    }
                }

                if (typeof filter.category != 'undefined' && filter.category != null)
                {
                    var tempListData: any[] = [];

                    for (var i = 0; i < filteredData.listData.length; i++)
                    {
                        if (filteredData.listData[i].productSubCategory.toString().trim().toLowerCase() === filter.category.trim().toLowerCase())
                        {
                            tempListData.push(filteredData.listData[i]);
                        }
                    }

                    filteredData.listData = tempListData;

                    for (var i = 0; i < filteredData.listData.length; i++)
                    {
                        if (filteredData.types.indexOf(filteredData.listData[i].type) === -1)
                        {
                            filteredData.types.push(filteredData.listData[i].type);
                        }
                    }
                }

                if (typeof filter.type != 'undefined' && filter.type !== null)
                {
                    tempListData = [];

                    for (var i = 0; i < filteredData.listData.length; i++)
                    {
                        if (filteredData.listData[i].type.toString().trim().toLowerCase() === filter.type.trim().toLowerCase())
                        {
                            tempListData.push(filteredData.listData[i]);
                        }
                    }

                    filteredData.listData = tempListData;
                }
            } 

            return filteredData;
        };
    }
} 

ko.utils.extendObservable = (target: Object, source: Object): void =>
{
    var prop: any, srcVal: any, isObservable: boolean = false;

    for (prop in source) {

        if (!source.hasOwnProperty(prop)) {
            continue;
        }

        if (ko.isWriteableObservable(source[prop])) {
            isObservable = true;
            srcVal = source[prop]();
        } else if (typeof (source[prop]) !== 'function') {
            srcVal = source[prop];
        }

        if (ko.isWriteableObservable(target[prop])) {
            target[prop](srcVal);
        } else if (target[prop] === null || target[prop] === undefined) {

            target[prop] = isObservable ? ko.observable(srcVal) : srcVal;

        } else if (typeof (target[prop]) !== 'function') {
            target[prop] = srcVal;
        }

        isObservable = false;
    }
};

ko.utils.clone = (obj: Object, emptyObj: Object): void =>
{
    var json = ko.toJSON(obj);
    var js = JSON.parse(json);

    ko.utils.extendObservable(emptyObj, js);
};