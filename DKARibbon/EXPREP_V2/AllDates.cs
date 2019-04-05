using DKAExcelStuff;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;

namespace EXPREP_V2
{
    public class AllDates
    {
        Master m;
        private readonly List<AllDates> _datesToUpdate;

        public AllDates() { }

        public AllDates(Master master)
        {
            m = master;
            _datesToUpdate = new List<AllDates>();
        }

        public int RowToUpdate { get; set; }
        public DateTime Received { get; set; }
        public DateTime RevisedScheduledDeliveryDate { get; set; }
        public DateTime POCreated { get; set; }
        public DateTime OriginalScheduledDelivery { get; set; }

        public bool IsDatesToUpdateInExpediteReport => _datesToUpdate.Count > 0 ? true : false;

        public int Year => POCreated.Year;
        public int Month => POCreated.Month;
        public string Quarter => KAXL.Quarter(POCreated);

        public int QReceivedDatesToUpdate { get; set; }
        public int QRevisedScheduledDeliveryDatesToUpdate { get; set; }

        public AllDates this[int i]
        {
            get => _datesToUpdate[i];
            set => _datesToUpdate[i] = value;
        }

        public void AddDateToExpRepUpdateList(int row, DateTime revisedSchedDeliveryDate, bool updateReceivedDate = false)
        {
            if (updateReceivedDate)
            {
                _datesToUpdate.Add(new AllDates
                {
                    RowToUpdate = row,
                    Received = ActualReceivedDate(),
                });
                QReceivedDatesToUpdate++;
            }
            else if(revisedSchedDeliveryDate != DateTime.MinValue)
            {
                _datesToUpdate.Add(new AllDates
                {
                    RowToUpdate = row,
                    RevisedScheduledDeliveryDate = revisedSchedDeliveryDate,
                });
                QRevisedScheduledDeliveryDatesToUpdate++;
            }            
        }

        private DateTime ActualReceivedDate() => DateTime.Today.DayOfWeek != DayOfWeek.Monday ?
            DateTime.Today.AddDays(-1) : DateTime.Today.AddDays(-3);

        public int QDatesToUpdate() => _datesToUpdate.Count;

        public void UpdateDatesOnExpediteReport()
        {
            m.kaxlApp.WS = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];

            WS ws = m.kaxlApp.WS;
            var k = m.kaxlApp.KAXL_RG;
            k = new KAXLApp.KAXLRange(m.kaxlApp, RangeType.WorkSheet);

            m.kaxlApp.ErrorTracker.ProgramStage = "Updating Dates In Expedite Report";

            for (int i = 0; i < QDatesToUpdate(); i++)
            {
                AllDates updateDates = _datesToUpdate[i];
                if (updateDates.Received != null)
                    ws.Cells[updateDates.RowToUpdate, m.ExpRepColumn.RecDate].Value = updateDates.Received;
                if (updateDates.RevisedScheduledDeliveryDate != null || updateDates.RevisedScheduledDeliveryDate != DateTime.MinValue)
                    ws.Cells[updateDates.RowToUpdate, m.ExpRepColumn.RevisedSchedDelDate].Value = updateDates.RevisedScheduledDeliveryDate;
            }
        }
    }
}
