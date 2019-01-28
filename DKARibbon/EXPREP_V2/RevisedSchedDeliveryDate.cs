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
    public class RevisedSchedDeliveryDate : Dates
    {
        private Master M;

        public RevisedSchedDeliveryDate(Master m) => M = m;

        public RevisedSchedDeliveryDate(string revDate) => MostRecentShedDeliveryDate = Convert.ToDateTime(revDate);

        public DateTime MostRecentShedDeliveryDate { get; set; }
        public int RowToUpdate { get; set; }

        public RevisedSchedDeliveryDate(int row, DateTime revDate)
        {
            RowToUpdate = row;
            MostRecentShedDeliveryDate = revDate;
        }
    }
    public class RevisedSchedDelDatesToUpdate
    {
        Master M;
        List<RevisedSchedDeliveryDate> revSchedDelDatesToUpdateList;

        public RevisedSchedDelDatesToUpdate(Master m)
        {
            M = m;
            revSchedDelDatesToUpdateList = new List<RevisedSchedDeliveryDate>();
        }
        public RevisedSchedDeliveryDate this[int i]
        {
            get => revSchedDelDatesToUpdateList[i];
            set => revSchedDelDatesToUpdateList[i] = value;
        }
        public void AddToUpdateList(int row, DateTime revDate)
        {
            // because datetime's can't be null, revDate set to Datetime.MinValue when blank - but this is filtered-out
            if (revDate != DateTime.MinValue)
            {
                revSchedDelDatesToUpdateList.Add(new RevisedSchedDeliveryDate(row, revDate));
            }
        }
        public int Q => revSchedDelDatesToUpdateList.Count;
        public bool IsRevisedSchedDelDatestoUpdate => Q > 0 ? true : false;

        public List<RevisedSchedDeliveryDate> GetRevisedSchedDeliveryDatesToUpdate() => revSchedDelDatesToUpdateList;

        public void UpdateRevisedDateOnExpRep()
        {
            // update exp rep with revised dates
            int Q = M.RevisedSchedDelDatesToUpdate.Q;
            int col = M.ExpRepColumn.RevisedSchedDelDate;
            WS ws = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            RevisedSchedDeliveryDate date;

            for (int i = 0; i < Q; i++)
            {                
                date = M.RevisedSchedDelDatesToUpdate[i];
                var c = ws.Cells[date.RowToUpdate, col];
                c.Value2 = date.MostRecentShedDeliveryDate;
                c.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            }
        }        
    }
}
