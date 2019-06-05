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

        public enum CleanStatusE { Open,Closed,Draft,Canceled,Received } // Received will be set to closed after update of rec dates                
        
        // for POLine Level
        public Status(string lineStatus, Master m, string poNum, string lineNumber, string approvalStatus)
        {
            CleanStatus = FormatStatus(lineStatus, m, poNum, lineNumber, approvalStatus);
        }

        public string ExpRepStatus { get; set; } // status of the line - if it is already in the Exp Report
        public CleanStatusE CleanStatus { get; set; } // final scrubbed status
           
        private CleanStatusE FormatStatus(string s, Master m, string poNum, string lineNumber, string approvalStatus)
        {
            string statusInExpRep, statusInAllPORep;
            try
            {
                string key = poNum + lineNumber;
                var po = m.PODictionaryInExpRep[key];
                statusInExpRep = po?.Status.ExpRepStatus;                
            }
            catch
            {
                statusInExpRep = s;
                statusInAllPORep = s;
            }
            string statusFromOpenLinesRep = s;

            if (approvalStatus == "Draft" || approvalStatus == "In review")
                return CleanStatusE.Draft;
            else if (statusFromOpenLinesRep == "Open order")
                return CleanStatusE.Open;
            else if (statusFromOpenLinesRep == "Received")
                return CleanStatusE.Received;
            else if (statusFromOpenLinesRep == "Cancelled")
                return CleanStatusE.Canceled;
            else if (statusFromOpenLinesRep == "Invoiced" || approvalStatus == "Finalized")
                return CleanStatusE.Closed;
            else
                return CleanStatusE.Open;
        }
    }
}
