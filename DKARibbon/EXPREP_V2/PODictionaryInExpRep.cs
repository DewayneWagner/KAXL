using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using DKAExcelStuff;

namespace EXPREP_V2
{
    public class PODictionaryInExpRep
    {
        public PODictionaryInExpRep() { }

        Master M;
        private readonly Dictionary<string, PODictionaryInExpRep> poLinesInDict;

        public PODictionaryInExpRep(string status, string poNum, double poLineNum, DateTime revDate, int expRepXLLineNum, 
            bool isReceivedDatePresent)
        {
            Status = new Status();
            Status.ExpRepStatus = status;
            PONum = poNum;
            POLineNum = poLineNum;
            MostRecRevDate = revDate;
            ExpRepXLLineNum = expRepXLLineNum;
            IsReceivedDatePresent = isReceivedDatePresent;
        }
        
        public string PONum { get; set; }
        public double POLineNum { get; set; }
        public Status Status { get; set; }
        public int ExpRepXLLineNum { get; set; }
        public DateTime MostRecRevDate { get; set; }
        public bool IsReceivedDatePresent { get; set; }

        public PODictionaryInExpRep(Master m)
        {
            M = m;
            poLinesInDict = new Dictionary<string, PODictionaryInExpRep>();

            WS ws = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            int LR = KAXL.LastRow(ws, 1) + 1;

            string status, poNum, key, recDateString;
            double poLineNum;
            DateTime revDate;
            int expRepXLLineNum;
            bool isRecDatePresent;
            string itemNum, itemDesc;

            for (int iRow = 2; iRow < LR; iRow++)
            {
                // for testing, to determine row which is erroring out
                //ws.Cells[1,1].Value2 = iRow;

                status = ws.Cells[iRow, M.ExpRepColumn.Status].Value2;
                poNum = ws.Cells[iRow, M.ExpRepColumn.PONumber].Value2;
                poLineNum = Math.Floor(ws.Cells[iRow, M.ExpRepColumn.LineNumber].Value2);
                
                var recDate = ws.Cells[iRow, M.ExpRepColumn.RecDate].Value2;
                recDateString = (recDate is string) ? recDate : Convert.ToString(recDate);
                isRecDatePresent = (recDateString is null) ? false : true;

                var rd = ws.Cells[iRow, M.ExpRepColumn.RevisedSchedDelDate].Value;
                revDate = (rd is DateTime) ? rd : Convert.ToDateTime(rd);
                expRepXLLineNum = iRow;

                // check for missing item descriptions & categories
                var itemNumber = ws.Cells[iRow, M.ExpRepColumn.ItemNumber].Value;
                itemDesc = ws.Cells[iRow, M.ExpRepColumn.ItemDescription].Value;
                itemNum = (itemNumber is string) ? itemNumber : Convert.ToString(itemNumber);
                
                if(itemNumber != null && itemDesc == null)
                {
                    // do this if there is an item number, but not description
                    Item i = M.ItemDict[itemNum];
                    ws.Cells[iRow, M.ExpRepColumn.ItemDescription].Value2 = i.Desc;
                }

                // check if vendor name is present, and add to update list if not.
                string vendorName = ws.Cells[iRow, M.ExpRepColumn.VendorName].Value2;
                if(vendorName == null)
                {
                    Vendor v = M.VendorDict[vendorName];
                    ws.Cells[iRow, M.ExpRepColumn.VendorName].Value2 = v.Name;
                }

                PODictionaryInExpRep po = new PODictionaryInExpRep(status, poNum, poLineNum, revDate, 
                    expRepXLLineNum, isRecDatePresent);

                key = poNum + poLineNum;

                if (!poLinesInDict.ContainsKey(key))
                {
                    poLinesInDict.Add(key, po);
                }
            }            
        }
        //public bool IsDuplicate(string poNum, double lineNum) => poLinesInDict.ContainsKey(poNum + lineNum);        
        public bool IsDuplicate(string poNum, double lineNum) => 
            (poLinesInDict.ContainsKey(poNum + Convert.ToString(Math.Floor(lineNum))))? true : false;            
        
        public PODictionaryInExpRep this[string key]
        {            
            get => key != null && poLinesInDict.ContainsKey(key) || 
                poLinesInDict.ContainsKey(Convert.ToString(Math.Floor(Convert.ToDouble(key))))
                ? poLinesInDict[key] : null;
            set => poLinesInDict[key] = value;
        }
    }
}
