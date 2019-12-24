using DKAExcelStuff;
using EXPREP_V2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using XL = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using static EXPREP_V2.Master;

namespace DKARibbon.SQLite_DataBase
{
    class DBUpdate
    {
        private KAXLApp k;
        private Dictionary<string, Item> _itemDictionary;

        public DBUpdate()
        {
            string path = @"R:\Supply Chain\ExpediteReport";
            string fileName = @"ExpediteReport.xlsx";
            string fullPath = Path.Combine(path, fileName);

            k = new KAXLApp(fullPath,(int)SheetNamesE.ExpRep);

            _itemDictionary = new Dictionary<string, Item>();

            UpdateItemDB();
        }
        
        private void UpdateItemDB()
        {
            k.WS = k.WB.Sheets[(int)SheetNamesE.MasterData];
            int firstRow = 2;
            int lastRow = KAXL.LastRow(k.WS, (int)MasterDataColumnsE.ItemNum);

            int firstCol = (int)MasterDataColumnsE.ItemNum;
            int lastCol = (int)MasterDataColumnsE.ItemCat;

            k.RG = k.WS.Range[k.WS.Cells[firstRow, firstCol], k.WS.Cells[lastRow, lastCol]];
            k.KAXL_RG = new KAXLApp.KAXLRange(k, RangeType.CodedRangeSetKAXLAppRG);

            for (int row = 1; row < k.KAXL_RG.Row.End; row++)
            {
                Item i = new Item();
                i.Num = Convert.ToString(k.KAXL_RG[row, 1]);
                i.Desc = Convert.ToString(k.KAXL_RG[row, 2]);
                i.Cat = Convert.ToString(k.KAXL_RG[row, 3]);

                if (!_itemDictionary.ContainsKey(i.Num))
                {
                    _itemDictionary.Add(i.Num, i);
                }
            }
        }
    }
}
