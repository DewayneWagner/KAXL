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
        private object[,] _scrubbedPOLinesObjArray;
        private List<ScrubbedPOLine> _scrubbedPOList;
        Master m;
        private int rowQ, colQ;

        public WriteObjectArrayToExpRep(Master master)
        {
            m = master;
            _scrubbedPOList = m.POLinesList.GetScrubbedPOLineList();
            
            rowQ = _scrubbedPOList.Count;
            colQ = m.ExpRepColumn.totalColumnsInExpRep;
            
            _scrubbedPOLinesObjArray = new object[rowQ, colQ];

            LoadObjArray();
            WriteArrayInExpRep();
            UpdateAdditionalInformationOnExpRep();
        }

        public object this[int r, int c]
        {
            get => _scrubbedPOLinesObjArray[r,c];
            set => _scrubbedPOLinesObjArray[r, c] = value;
        }

        private void LoadObjArray()
        {
            for (int i = 0; i < rowQ; i++)
            {
                //i--; // to account for 0 index in list, but not in array
                ScrubbedPOLine po = _scrubbedPOList[i];

                i++;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.AttentionInfo] = po.Source.OriginalAttentionInfo;                
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Category] = po.Category.CleanCategory;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Createdby] = po.Source.CreatedBy;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Curr] = po.Cash.Currency;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.DateAdded] = DateTime.Today;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Direct] = po.Direct;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Entity] = po.Entity;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ExpediteRequired] = po.ExpediteRequired;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Expeditor] = po.Source.CreatedBy;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ICO] = po.ICO;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ItemDescription] = po.ItemX.Desc;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ItemNumber] = po.ItemX.Num;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.LineNumber] = po.LineNumber;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Month] = po.Dates.Month;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.NetAmount] = po.Cash.NetAmount;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.OriginalSchedDelDate] = po.Dates.OriginalScheduledDelivery;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.POCreatedDate] = po.Dates.POCreated;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.PONumber] = po.PONum;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.POSourceCode] = po.Source.Code;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.POSourceType] = po.Source.Type;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Quantity] = po.Quantity;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Quarter] = po.Dates.Quarter;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.RecDate] = po.Dates.Received;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Receiver] = po.Receiver;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Requester] = po.Source.Requester;

                if(po.Dates.RevisedScheduledDeliveryDate == DateTime.MinValue)
                {
                    _scrubbedPOLinesObjArray[i, m.ExpRepColumn.RevisedSchedDelDate] = po.Dates.RevisedScheduledDeliveryDate;
                }

                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Status] = po.Status.CleanStatus;                
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.UnitPriceUSD] = po.Cash.UnitPriceUSD;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.USD] = po.Cash.USD;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.VendorAccount] = po.Vendor.Code;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.VendorName] = po.Vendor.Name;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.WH] = po.WH;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Year] = po.Dates.Year;
            }
        }
        public void WriteArrayInExpRep()
        {
            WS ws = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            
            int firstRowToWriteArrayAfterData = KAXL.LastRow(ws, 1) + 1;
            int lastRowToWriteArray = firstRowToWriteArrayAfterData + _scrubbedPOLinesObjArray.GetLength(0);
            int totalColumns = _scrubbedPOLinesObjArray.GetLength(1);

            RG rg = ws.Range[ws.Cells[firstRowToWriteArrayAfterData, 1], ws.Cells[lastRowToWriteArray, totalColumns + 1]];

            rg.Value = _scrubbedPOLinesObjArray;
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
