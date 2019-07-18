using DKAExcelStuff;
using EXPREP_V2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DKAExcelStuff.KAXLApp;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using Microsoft.Office.Interop.Excel;

namespace DKARibbon.BOM_Scrubber
{
    public class ExpRepData
    {
        private enum RequiredFields { Null, ItemNum, PODate, PONumber, VEndorName, UnitCost, Total }
        private List<int> _reqFieldsColNumbers;
        private Dictionary<string, ItemLastPurchased> _purchasedItemsDictionary;
  
        public ExpRepData()
        {
            kaxlApp = new KAXLApp(@"R:\Supply Chain\ExpediteReport\ExpediteReport.xlsx");
            
            Col = new ExpRepColumn(GetColumnHeadingsList());
            _reqFieldsColNumbers = GetListOfColNumbers();

            LoadDictionaryOfItemsPurchased(ItemLastPurchased.LoadList(kaxlApp, Col));
        }

        private List<int> GetListOfColNumbers()
        {
            List<int> _colNums = new List<int>()
            {
                Col.ItemNumber,
                Col.POCreatedDate,
                Col.PONumber,
                Col.VendorName,
                Col.UnitPriceUSD,
            };
            return _colNums;
        }
        private List<string> GetColumnHeadingsList()
        {
            List<string> _colHeadingsLists = new List<string>();
            var k = kaxlApp.KAXL_RG;

            for (int c = 1; c < k.Col.Q; c++)
            {
                _colHeadingsLists.Add(Convert.ToString(k[1, c]));
            }

            return _colHeadingsLists;
        }


        public ExpRepColumn Col { get; }
        KAXLApp kaxlApp { get; }

        private void LoadDictionaryOfItemsPurchased(List<ItemLastPurchased> _allItems)
        {
            _purchasedItemsDictionary = new Dictionary<string, ItemLastPurchased>();
            var _sortedList = _allItems.OrderByDescending(d => d.PODate);

            foreach (ItemLastPurchased item in _sortedList)
            {
                if (!_purchasedItemsDictionary.ContainsKey(item.ItemNum))
                {
                    _purchasedItemsDictionary.Add(item.ItemNum, item);
                }
            }
        }

        private class ItemLastPurchased
        {
            public string ItemNum { get; set; }
            public DateTime PODate { get; set; }
            public string PONumber { get; set; }
            public string VendorName { get; set; }
            public double UnitCost { get; set; }

            internal static List<ItemLastPurchased> LoadList(KAXLApp kaxlApp, ExpRepColumn c)
            {
                var k = kaxlApp.KAXL_RG;
                double unitCost;
                string itemNum;

                List<ItemLastPurchased> _itemLastPurchasedList = new List<ItemLastPurchased>();

                for (int row = 2; row < k.Row.Q; row++)
                {
                    itemNum = Convert.ToString(k[row, c.ItemNumber]);

                    if(itemNum != null)
                    {
                        _itemLastPurchasedList.Add(new ItemLastPurchased()
                        {
                            ItemNum = itemNum,
                            PODate = KAXL.ReadDateTime(k[row, c.POCreatedDate]),
                            PONumber = Convert.ToString(k[row, c.PONumber]),
                            UnitCost = double.TryParse(Convert.ToString(k[row,c.UnitPriceUSD]),out unitCost)? unitCost : 0,
                            VendorName = Convert.ToString(k[row, c.VendorName])
                        });
                    }                    
                }
                return _itemLastPurchasedList;
            }
        }
    }
}
