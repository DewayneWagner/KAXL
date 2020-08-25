using System;
using System.Collections.Generic;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using DKAExcelStuff;
using MDC = EXPREP_V2.Master.MasterDataColumnsE;

namespace EXPREP_V2
{
    public class Item
    {
        private enum ItemColumnOrder { Nada,Num,Desc,Cat}
        Master M;
        public Item(Master m)
        {
            M = m;
            _itemDictionary = new Dictionary<string, Item>();
            _itemNumbersThatArentInDictL = new List<string>();
            _itemsThatAreMissingDescriptionAndCategoryInExpRep = new List<Item>();
            LoadItemDictionary();
        }

        private readonly Dictionary<string, Item> _itemDictionary;
        private List<string> _itemNumbersThatArentInDictL;
        private List<Item> _itemsThatAreMissingDescriptionAndCategoryInExpRep;

        public Item() {}
        
        public Item(string num, string desc, string cat)
        {
            Num = num;
            Desc = desc;
            Cat = cat;
        }
        public string Num { get; set; }
        public string Desc { get; set; }
        public string Cat { get; set; }
        public int QItemsInBOM { get; set; }
        public int ExpRepRow { get; set; }

        private void LoadItemDictionary()
        {
            M.kaxlApp.WS = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.MasterData];

            WS ws = M.kaxlApp.WS;
            var k = M.kaxlApp.KAXL_RG;

            M.kaxlApp.RG = ws.Range[ws.Cells[KAXL.FindFirstRowAfterHeader(ws), (int)MDC.ItemNum], ws.Cells[KAXL.LastRow(ws, (int)MDC.ItemNum), (int)MDC.ItemCat]];
            k = new KAXLApp.KAXLRange(M.kaxlApp, RangeType.CodedRangeSetKAXLAppRG);

            string itemNum;

            for (int r = 1; r < k.Row.End; r++)
            {
                itemNum = Convert.ToString(k[r, (int)ItemColumnOrder.Num]);

                if (_itemDictionary.ContainsKey(itemNum) || itemNum == null)
                {
                    itemNum = null;
                }
                else
                {
                    _itemDictionary.Add(itemNum, new Item()
                    {
                        Num = itemNum,
                        Desc = (string)k[r, (int)ItemColumnOrder.Desc],
                        Cat = (string)k[r, (int)ItemColumnOrder.Cat]
                    });
                }
            }
        }
        public Item this[string key]
        {
            get => key != null && _itemDictionary.ContainsKey(key) && key.Length > 5 ?
                _itemDictionary[key] : AddItemToItemNumbersThatArentInDictL(key);
            ///set => itemDict[key] = value;
        }
        public Item AddItemToItemNumbersThatArentInDictL(string key)
        {
            if (key != null && key.Length > 5)
                _itemNumbersThatArentInDictL.Add(key);

            string num = key != null ? key : null;
            string desc = null;
            string cat = null;

            return new Item(num, desc, cat);
        }
        public List<string> GetItemNumbersThatArentInDictList() => _itemNumbersThatArentInDictL;        
        public bool IsItemsThatArentInDict() => _itemNumbersThatArentInDictL.Count > 0 ? true : false;
        public void AddItemsInExpRepMissingDescriptions(object itemNum, int rowInExpRep)
        {
            string num = Convert.ToString(itemNum);

            if (_itemDictionary.ContainsKey(num))
            {
                Item i = _itemDictionary[num];
                i.ExpRepRow = rowInExpRep;
                _itemsThatAreMissingDescriptionAndCategoryInExpRep.Add(i);
            }
        }
        public bool IsItemsThatNeedToHaveDescriptionsUpdated() => _itemsThatAreMissingDescriptionAndCategoryInExpRep.Count > 0 ? true : false;
        public List<Item> GetListOfItemsThatNeedToHaveItemDescriptionsUpdated() => _itemsThatAreMissingDescriptionAndCategoryInExpRep;
    }
}
