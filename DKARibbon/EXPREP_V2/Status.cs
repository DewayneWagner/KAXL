using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using Microsoft.Office.Interop.Excel;
using DKAExcelStuff;

namespace EXPREP_V2
{
    public class Status
    {
        public Status() { }

        private Dictionary<string, Status> statusDict;
        private Master M;
        enum StatusColNums { Nada,PONum,POStatus, LineStatus }
        enum CleanStatusE { Open,Closed,Draft,Canceled,Received } // Received will be set to closed after update of rec dates

        List<string> AllStatusL = new List<string>() { "Invoiced", "Received", "Open order", "Canceled" };
        enum AllStatusE { Invoiced,Received,OpenOrder,Canceled }

        public Status(Master m)
        {
            M = m;
            statusDict = new Dictionary<string, Status>();
            LoadDict();
        }
        // for the AllPOs list
        public Status(string poNum, string po, string line)
        {
            PONum = poNum;
            PO = po;
            Line = line;
        }
        // for POLine Level
        public Status(string lineStatus, Master m, string poNum, string lineNumber)
        {
            CleanStatus = FormatStatus(lineStatus, m, poNum, lineNumber);
        }

        //public Master Master { get; set; }
        public string PONum { get; set; }
        public string PO { get; set; } // status of the PO from All PO report
        public string Line { get; set; } // status of the PO Line from Open Lines Report
        public string ExpRepStatus { get; set; } // status of the line - if it is already in the Exp Report
        public string CleanStatus { get; set; } // final scrubbed status
                
        private void LoadDict()
        {
            WS ws = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.AllPOs];
            int LR = KAXL.LastRow(ws, 1);
            string poNum, po, line;                
            
            for (int iRow = 2; iRow <= LR; iRow++)
            {
                poNum = ws.Cells[iRow, (int)StatusColNums.PONum].Value2;
                po = ws.Cells[iRow, (int)StatusColNums.POStatus].Value2;
                line = ws.Cells[iRow, (int)StatusColNums.LineStatus].Value2;

                if(!statusDict.ContainsKey(poNum))
                    statusDict.Add(poNum, (new Status(poNum,po,line)));
            }
            ws.Cells.ClearContents();
        }
        public Status this[string poNum] => statusDict.ContainsKey(poNum) || poNum != null ? statusDict[poNum] : null;
        
        private string FormatStatus(string s, Master m, string poNum, string lineNumber)
        {
            string statusInExpRep, statusInAllPORep;
            try
            {
                string key = poNum + lineNumber;
                var po = m.PODictionaryInExpRep[key];

                statusInExpRep = po?.Status.ExpRepStatus;
                var allpo = m.AllPOsDict[poNum];
                statusInAllPORep = allpo?.PO;
            }
            catch
            {
                statusInExpRep = s;
                statusInAllPORep = s;
            }
            string statusFromOpenLinesRep = s;

            if (statusInAllPORep == "Draft" || statusInAllPORep == "In review")
                return Convert.ToString((CleanStatusE)(int)CleanStatusE.Draft);
            else if (statusFromOpenLinesRep == "Open order")
                return Convert.ToString((CleanStatusE)(int)CleanStatusE.Open);
            else if (statusFromOpenLinesRep == "Received")
                return Convert.ToString((CleanStatusE)(int)CleanStatusE.Received);
            else if (statusFromOpenLinesRep == "Cancelled")
                return Convert.ToString((CleanStatusE)(int)CleanStatusE.Canceled);
            else if (statusFromOpenLinesRep == "Invoiced" || statusInAllPORep == "Finalized")
                return Convert.ToString((CleanStatusE)(int)CleanStatusE.Closed);
            else
                return Convert.ToString((CleanStatusE)(int)CleanStatusE.Open);
        }
        public bool IsInAllPOReport(string key) => statusDict.ContainsKey(key);
    }
}
