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
    public class ReceivedDate : Dates
    {
        public ReceivedDate() { }
        
        private Master M;

        public ReceivedDate(Master m) => M = m;

        public ReceivedDate(int row)
        {
            RowToUpdate = row;            
        }        
        
        public int RowToUpdate { get; set; }

        public DateTime Actual => DateTime.Today.DayOfWeek != DayOfWeek.Monday ? 
            DateTime.Today.AddDays(-1) : DateTime.Today.AddDays(-3);
    }
    public class ReceivedDateList : ReceivedDate
    {        
        private List<ReceivedDate> recDatesToUpdate;
        Master M;

        public ReceivedDateList(Master m)
        {
            M = m;
            recDatesToUpdate = new List<ReceivedDate>();
        }

        public void AddToUpdateList(int row) => recDatesToUpdate.Add(new ReceivedDate(row));

        public ReceivedDate this[int i]
        {
            get => recDatesToUpdate[i];
            set => recDatesToUpdate[i] = value;
        }
        public int Q => recDatesToUpdate.Count;

        public List<ReceivedDate> GetReceivedDatesToUpdateList() => recDatesToUpdate;
        public bool IsReceivedDateItemsToUpdate => Q > 0 ? true : false;

        public void UpdateReceivedDatesOnExpRep()
        {
            // update exp rep with received dates
            int Q = M.ReceivedDateList.Q;
            int col = M.ExpRepColumn.RecDate;
            WS ws = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            ReceivedDate date;

            for (int i = 0; i < Q; i++)
            {
                date = M.ReceivedDateList[i];
                ws.Cells[date.RowToUpdate, col].Value = date.Actual;
                ws.Cells[date.RowToUpdate, M.ExpRepColumn.Status].Value2 = "Closed";
            }
        }
    }
}
