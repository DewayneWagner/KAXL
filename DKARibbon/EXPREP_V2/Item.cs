using System;
using System.Collections.Generic;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using DKAExcelStuff;

namespace EXPREP_V2
{
    public class Item
    {
        Master M;
        public Item(Master m) => M = m;

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
    }
    public class ItemDict
    {
        private readonly Dictionary<string,Item> itemDict;
        private List<string> itemNumbersThatArentInDictL;

        Master M;
        public ItemDict(Master m)
        {
            M = m;
            itemDict = new Dictionary<string, Item>();
            itemNumbersThatArentInDictL = new List<string>();
            LoadDict();
        }

        public ItemDict() { }

        public Item this[string key]
        {
            get => key != null && itemDict.ContainsKey(key) && key.Length > 5 ? 
                itemDict[key] : AddItemToItemNumbersThatArentInDictL(key);
            set => itemDict[key] = value;
        }
        public Item AddItemToItemNumbersThatArentInDictL(string key)
        {
            if(key!=null && key.Length > 5)
                itemNumbersThatArentInDictL.Add(key);

            string num = key != null ? key : null;
            string desc = null;
            string cat = null;

            return new Item(num, desc, cat);
        }
        private void LoadDict()
        {
            int LR, startRow = 2, numC = 12, descC = 13, catC = 14;

            WS ws = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.MasterData];
            LR = KAXL.LastRow(ws, numC);
            string itemNum, itemDesc, itemCat;

            for (int j = startRow; j <= LR; j++)
            {
                var val = ws.Cells[j, numC].Value2;
                itemNum = (val is string) ? val : Convert.ToString(val);
                itemDesc = ws.Cells[j, descC].Value2;
                itemCat = ws.Cells[j, catC].Value2;

                Item i = new Item(itemNum, itemDesc, itemCat);

                if(itemNum != null && !itemDict.ContainsKey(itemNum))
                    itemDict.Add(itemNum,i);
            }
        }
        public List<string> GetItemNumbersThatArentInDictList() => itemNumbersThatArentInDictL;
        public bool IsItemsThatArentInDict() => (itemNumbersThatArentInDictL.Count > 0) ? true : false;

        public Item GetItem(string key)
        {
            if (itemDict.ContainsKey(key)) 
            {
                return itemDict[key];
            }
            else
            {
                itemNumbersThatArentInDictL.Add(key);
                return null;
            }
        }        
    }
}
