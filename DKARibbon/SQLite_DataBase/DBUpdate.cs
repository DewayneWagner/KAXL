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
        private Dictionary<string, POLineDB> _poLinesInExpRep;
        private Dictionary<string, Item> _itemDictionary;
                
        public DBUpdate()
        {
            string path = @"R:\Supply Chain\ExpediteReport";
            string fileName = @"ExpediteReport.xlsx";
            string fullPath = Path.Combine(path, fileName);

            k = new KAXLApp(fullPath,(int)SheetNamesE.ExpRep);

            _poLinesInExpRep = new Dictionary<string, POLineDB>();
            _itemDictionary = new Dictionary<string, Item>();

            UpdatePOLineDB();
            UpdateItemDB();
        }
        public void UpdatePOLineDB()
        {
            int firstRow = 2;
            int lastRow = KAXL.LastRow(k.WS, 1);

            int firstCol = 1;
            int lastCol = KAXL.LastCol(k.WS, 2);

            k.RG = k.WS.Range[k.WS.Cells[firstRow, firstCol], k.WS.Cells[lastRow, lastCol]];
            k.KAXL_RG = new KAXLApp.KAXLRange(k, RangeType.CodedRangeSetKAXLAppRG);

            List<string> colHeadings = new List<string>();
            colHeadings.Add("Nada");

            for (int column = 1; column <= lastCol; column++)
            {
                colHeadings.Add(Convert.ToString(k.KAXL_RG[1, column]));
            }

            ExpRepColumn sourceColID = new ExpRepColumn(colHeadings);
            string key;

            string poNumber;
            double unitPrice;
            bool isICO;
            string itemNum;
            double lineNumber;
            DateTime mostRecentlyScheduledDeliveryDate;
            Status status;
            DateTime poCreatedDate;
            double quantity;
            string vendorName;

            for (int row = 2; row < k.KAXL_RG.Row.End; row++)
            {
                //poNumber = Convert.ToString(k.KAXL_RG[row, sourceColID.PONumber]);
                //unitPrice = Convert.ToDouble(k.KAXL_RG[row, sourceColID.UnitPriceUSD]);
                //isICO = ParseBool(k.KAXL_RG[row, sourceColID.ICO]);
                //itemNum = Convert.ToString(k.KAXL_RG[row, sourceColID.ItemNumber]);
                //lineNumber = Math.Round(Convert.ToDouble(k.KAXL_RG[row, sourceColID.LineNumber]), 1);
                //mostRecentlyScheduledDeliveryDate = KAXL.ReadDateTime(k.KAXL_RG[row, sourceColID.RevisedSchedDelDate]);
                //status = new Status() { ExpRepStatus = Convert.ToString(k.KAXL_RG[row, sourceColID.Status]) };
                //poCreatedDate = KAXL.ReadDateTime(k.KAXL_RG[row, sourceColID.POCreatedDate]);
                //quantity = Convert.ToDouble(k.KAXL_RG[row, sourceColID.Quantity]);
                //vendorName = Convert.ToString(k.KAXL_RG[row, sourceColID.VendorName]);

                //key = poNumber + Convert.ToString(lineNumber);

                //_poLinesInExpRep.Add(key,new POLineDB(poNumber,lineNumber)
                //{
                //    IsICO = isICO,
                //    ItemNum = itemNum,
                //    LineNumber = lineNumber,
                //    MostRecentlyScheduledDeliveryDate = mostRecentlyScheduledDeliveryDate,
                //    POCreatedDate = poCreatedDate,
                //    PONumber = poNumber,
                //    Quantity = quantity,
                //    Status = status.CleanStatus,
                //    UnitPrice = unitPrice,
                //    VendorName = vendorName,
                //});

                poNumber = Convert.ToString(k.KAXL_RG[row, sourceColID.PONumber]);
                lineNumber = Math.Round(Convert.ToDouble(k.KAXL_RG[row, sourceColID.LineNumber]), 1);
                status = new Status() { ExpRepStatus = Convert.ToString(k.KAXL_RG[row, sourceColID.Status]) };

                POLineDB po = new POLineDB(poNumber,lineNumber)
                {
                    //PONumber = Convert.ToString(k.KAXL_RG[row, sourceColID.PONumber]),
                    UnitPrice = Convert.ToDouble(k.KAXL_RG[row, sourceColID.UnitPriceUSD]),
                    IsICO = ParseBool(k.KAXL_RG[row, sourceColID.ICO]),
                    ItemNum = Convert.ToString(k.KAXL_RG[row, sourceColID.ItemNumber]),
                    //LineNumber = Convert.ToDouble(k.KAXL_RG[row, sourceColID.LineNumber]),
                    MostRecentlyScheduledDeliveryDate = KAXL.ReadDateTime(k.KAXL_RG[row, sourceColID.RevisedSchedDelDate]),
                    Status = status.CleanStatus,
                    POCreatedDate = KAXL.ReadDateTime(k.KAXL_RG[row, sourceColID.POCreatedDate]),
                    Quantity = Convert.ToDouble(k.KAXL_RG[row, sourceColID.Quantity]),
                    VendorName = Convert.ToString(k.KAXL_RG[row, sourceColID.VendorName])
                };
                key = po.PONumber + Convert.ToString(po.LineNumber);

                if (!_poLinesInExpRep.ContainsKey(key))
                {
                    _poLinesInExpRep.Add(key, po);
                } 
            }
            bool ParseBool(object ico) => (Convert.ToString(ico) == "true") ? true : false;
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
