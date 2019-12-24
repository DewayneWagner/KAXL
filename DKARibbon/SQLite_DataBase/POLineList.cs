using DKAExcelStuff;
using EXPREP_V2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DKARibbon.SQLite_DataBase
{
    class POLineList : List<POLineDB>
    {
        private KAXLApp k;
        public POLineList(KAXLApp kaxlApp)
        {
            k = kaxlApp;
            UpdatePOLineDB();
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
            
            string poNumber;
            double lineNumber;
            Status status;

            for (int row = 2; row < k.KAXL_RG.Row.End; row++)
            {
                poNumber = Convert.ToString(k.KAXL_RG[row, sourceColID.PONumber]);
                lineNumber = Math.Round(Convert.ToDouble(k.KAXL_RG[row, sourceColID.LineNumber]), 1);
                status = new Status() { ExpRepStatus = Convert.ToString(k.KAXL_RG[row, sourceColID.Status]) };

                POLineDB po = new POLineDB(poNumber, lineNumber)
                {
                    UnitPrice = Convert.ToDouble(k.KAXL_RG[row, sourceColID.UnitPriceUSD]),
                    IsICO = ParseBool(k.KAXL_RG[row, sourceColID.ICO]),
                    ItemNum = Convert.ToString(k.KAXL_RG[row, sourceColID.ItemNumber]),
                    MostRecentlyScheduledDeliveryDate = KAXL.ReadDateTime(k.KAXL_RG[row, sourceColID.RevisedSchedDelDate]),
                    Status = status.CleanStatus,
                    POCreatedDate = KAXL.ReadDateTime(k.KAXL_RG[row, sourceColID.POCreatedDate]),
                    Quantity = Convert.ToDouble(k.KAXL_RG[row, sourceColID.Quantity]),
                    VendorName = Convert.ToString(k.KAXL_RG[row, sourceColID.VendorName])
                };
                this.Add(po);
            }
            bool ParseBool(object ico) => (Convert.ToString(ico) == "true") ? true : false;
        }
    }
}
