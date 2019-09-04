using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using DKAExcelStuff;
using System.Runtime.InteropServices;

namespace EXPREP_V2
{
    public class WriteObjectArrayToExpRep
    {
        private object[,] _scrubbedPOLinesObjArrayOneIndexedColumns;
        private object[,] _scrubbedPOLinesObjZeroIndexedArrayZeroIndexedColumns;
        private List<ScrubbedPOLine> _scrubbedPOList;
        Master m;
        private int rowQ, colQ;

        public WriteObjectArrayToExpRep(Master master)
        {
            m = master;
            _scrubbedPOList = m.POLinesList.GetScrubbedPOLineList();
            
            rowQ = _scrubbedPOList.Count;
            colQ = m.ExpRepColumn.totalColumnsInExpRep;
            
            _scrubbedPOLinesObjArrayOneIndexedColumns = new object[rowQ, colQ];
            _scrubbedPOLinesObjZeroIndexedArrayZeroIndexedColumns = new object[rowQ, colQ - 1];

            LoadObjArray();
            WriteArrayInExpRep();
            UpdateAdditionalInformationOnExpRep();
        }
        private void LoadObjArray()
        {
            for (int i = 0; i < rowQ; i++)
            {
                ScrubbedPOLine po = _scrubbedPOList[i];

                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.AttentionInfo] = po.Source.OriginalAttentionInfo;                
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Category] = po.Category.CleanCategory;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Createdby] = po.Source.CreatedBy;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Curr] = po.Cash.Currency;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.DateAdded] = DateTime.Today;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Direct] = po.Direct;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Entity] = po.Entity;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.ExpediteRequired] = po.ExpediteRequired;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Expeditor] = po.Source.CreatedBy;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.ICO] = po.ICO;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.ItemDescription] = po.ItemX.Desc;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.ItemNumber] = po.ItemX.Num;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.LineNumber] = po.LineNumber;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Month] = po.Dates.Month;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.NetAmount] = po.Cash.NetAmount;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.OriginalSchedDelDate] = po.Dates.OriginalScheduledDelivery;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.POCreatedDate] = po.Dates.POCreated;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.PONumber] = po.PONum;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.POSourceCode] = po.Source.Code;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.POSourceType] = Convert.ToString((Source.SourceType)po.Source.Type);
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Quantity] = po.Quantity;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Quarter] = po.Dates.Quarter;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.IsRush] = po.Source.IsRush;

                if(po.Status.CleanStatus == Status.CleanStatusE.Received)
                {
                    _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.RecDate] = po.Dates.Received;
                    _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Status] = Convert.ToString(Status.CleanStatusE.Closed);
                }
                else
                {
                    _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Status] = Convert.ToString((Status.CleanStatusE)po.Status.CleanStatus);
                }

                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Receiver] = po.Receiver;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Requester] = po.Source.Requester;

                if(po.Dates.RevisedScheduledDeliveryDate != DateTime.MinValue)
                {
                    _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.RevisedSchedDelDate] = po.Dates.RevisedScheduledDeliveryDate;
                }
                
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.UnitPriceUSD] = po.Cash.UnitPriceUSD;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.USD] = po.Cash.USD;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.VendorAccount] = po.Vendor.Code;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.VendorName] = po.Vendor.Name;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.WH] = po.WH;
                _scrubbedPOLinesObjArrayOneIndexedColumns[i, m.ExpRepColumn.Year] = po.Dates.Year;
            }
            for (int r = 0; r < _scrubbedPOLinesObjArrayOneIndexedColumns.GetLength(0); r++)
            {
                for (int c = 1; c < _scrubbedPOLinesObjArrayOneIndexedColumns.GetLength(1); c++)
                {
                    _scrubbedPOLinesObjZeroIndexedArrayZeroIndexedColumns[r, c - 1] = _scrubbedPOLinesObjArrayOneIndexedColumns[r, c];
                }
            }
        }
        public void WriteArrayInExpRep()
        {
            WS ws = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            
            int firstRowToWriteArrayAfterData = KAXL.LastRow(ws, 1) + 1;
            int numberOfRows = _scrubbedPOLinesObjZeroIndexedArrayZeroIndexedColumns.GetLength(0);

            int lastRowToWriteArray = firstRowToWriteArrayAfterData + numberOfRows;
            int totalColumns = _scrubbedPOLinesObjZeroIndexedArrayZeroIndexedColumns.GetLength(1);

            RG rg = ws.Range[ws.Cells[firstRowToWriteArrayAfterData, 1], ws.Cells[lastRowToWriteArray-1, totalColumns]];

            rg.Value = _scrubbedPOLinesObjZeroIndexedArrayZeroIndexedColumns;

            m.updateMetrics.QTotalUpdatedLines = numberOfRows;
        }
        public void UpdateAdditionalInformationOnExpRep()
        {
            if (m.Dates.IsDatesToUpdateInExpediteReport)
            {
                m.kaxlApp.ErrorTracker.ProgramStage = "Updating Dates in Expedite Report";

                m.updateMetrics.QUpdatedReceivedDates = m.Dates.QReceivedDatesToUpdate;
                m.updateMetrics.QUpdatedRevisedDeliveryDates = m.Dates.QRevisedScheduledDeliveryDatesToUpdate;
                m.Dates.UpdateDatesOnExpediteReport();
            }

            WS expRep = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.MasterData];

            //Update vendor list with vendor numbers that aren't in dictionary
            if (m.VendorDict.IsVendorNumbersThatArentInDict())
            {
                m.kaxlApp.ErrorTracker.ProgramStage = "Updating vendor names in vendor list";

                int col = (int)Master.MasterDataColumnsE.VendorAccount;
                int NR = KAXL.LastRow(expRep, col) + 1;

                List<string> vendorNamesNotInDictionary = m.VendorDict.VendorNumbersThatArentInDictL();

                foreach (string VendorNumber in vendorNamesNotInDictionary)
                {
                    expRep.Cells[NR, col].Value2 = VendorNumber;
                    NR++;
                }
            }
            // Update Item List with item numbers not in dictionary
            if (m.ItemDict.IsItemsThatArentInDict())
            {
                m.kaxlApp.ErrorTracker.ProgramStage = "Updating Item's that aren't in dictionary";

                int col = (int)Master.MasterDataColumnsE.ItemNum;
                int NR = KAXL.LastRow(expRep, col) + 1;

                List<string> itemNumbersNotInDictionary = m.ItemDict.GetItemNumbersThatArentInDictList();

                foreach (string item in itemNumbersNotInDictionary)
                {
                    if (item != null)
                    {
                        expRep.Cells[NR, col].Value2 = item;
                        NR++;
                    }
                }
            }
            m.stopWatch.EndTime = DateTime.Now;
        }
    }
}
