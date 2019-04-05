using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using System.Reflection;
using System.Runtime.InteropServices;
using DKAExcelStuff;
using System.Windows.Forms;

namespace EXPREP_V2
{
    public class AddToExpRep
    {
        public AddToExpRep(Master M)
        {
            //Worksheet ws = M.kaxlApp.WB.Sheets[Convert.ToString((Master.SheetNamesE)((int)Master.SheetNamesE.ExpRep))];
            WS ws = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];

            if (ws.AutoFilter != null)
                ws.AutoFilterMode = false;

            int length = M.POLinesList.ListQ;

            var dCol = M.ExpRepColumn;
            //ExpRepColumn dCol = new ExpRepColumn(ws);
            int nextRow = KAXL.LastRow(ws,1) + 1;

            M.kaxlApp.ErrorTracker.ProgramStage = "Writing to Expedite Report";
            
            // load list with column headings
            for (int i = 1; i < length; i++)
            {
                M.kaxlApp.ErrorTracker.Row = i;

                ScrubbedPOLine po = M.POLinesList[i];

                if(po.Status.CleanStatus != Status.CleanStatusE.Canceled)
                {
                    // POSource Class
                    ws.Cells[nextRow, dCol.AttentionInfo].Value2 = po.Source.OriginalAttentionInfo;
                    ws.Cells[nextRow, dCol.POSourceType].Value2 = Convert.ToString(po.Source.Type);
                    ws.Cells[nextRow, dCol.POSourceCode].Value2 = po.Source.Code;
                    ws.Cells[nextRow, dCol.Requester].Value2 = po.Source.Requester;
                    ws.Cells[nextRow, dCol.Createdby].Value2 = po.Source.CreatedBy;
                    ws.Cells[nextRow, dCol.Expeditor].Value2 = po.Source.CreatedBy;

                    // Cash Class
                    ws.Cells[nextRow, dCol.CAD].Value2 = po.Cash.CAD;
                    ws.Cells[nextRow, dCol.Curr].Value2 = po.Cash.Currency;
                    ws.Cells[nextRow, dCol.UnitPriceCAD].Value2 = po.Cash.UnitPriceCAD;
                    ws.Cells[nextRow, dCol.UnitPriceUSD].Value2 = po.Cash.UnitPriceUSD;
                    ws.Cells[nextRow, dCol.NetAmount].Value2 = po.Cash.NetAmount;
                    ws.Cells[nextRow, dCol.USD].Value2 = po.Cash.USD;

                    // No Class
                    ws.Cells[nextRow, dCol.Entity].Value2 = po.Entity;
                    ws.Cells[nextRow, dCol.ExpediteRequired].Value2 = po.ExpediteRequired;
                    ws.Cells[nextRow, dCol.ICO].Value2 = po.ICO;
                    ws.Cells[nextRow, dCol.LineNumber].Value2 = po.LineNumber;
                    ws.Cells[nextRow, dCol.PONumber].Value2 = po.PONum;
                    ws.Cells[nextRow, dCol.Quantity].Value2 = po.Quantity;
                    ws.Cells[nextRow, dCol.Direct].Value2 = Convert.ToString(po.Direct);

                    Status.CleanStatusE status = (po.Status.CleanStatus == Status.CleanStatusE.Received) ? Status.CleanStatusE.Closed : po.Status.CleanStatus;
                    ws.Cells[nextRow, dCol.Status].Value2 = Convert.ToString(status);

                    ws.Cells[nextRow, dCol.WH].Value2 = po.WH;
                    ws.Cells[nextRow, dCol.Receiver].Value2 = po.Receiver;

                    // Item Class
                    ws.Cells[nextRow, dCol.ItemDescription].Value2 = po.ItemX.Desc;
                    ws.Cells[nextRow, dCol.ItemNumber].Value2 = po.ItemX.Num;

                    // Date Class                
                    ws.Cells[nextRow, dCol.Month].Value2 = po.Dates.Month;
                    ws.Cells[nextRow, dCol.OriginalSchedDelDate].Value = po.Dates.OriginalScheduledDelivery;
                    ws.Cells[nextRow, dCol.POCreatedDate].Value = po.Dates.POCreated;
                    ws.Cells[nextRow, dCol.Quarter].Value2 = po.Dates.Quarter;
                    ws.Cells[nextRow, dCol.Year].Value2 = po.Dates.Year;
                    ws.Cells[nextRow, dCol.DateAdded].Value = DateTime.Today;
                    
                    if(po.Dates.RevisedScheduledDeliveryDate != DateTime.MinValue)
                    {
                        ws.Cells[nextRow, dCol.RevisedSchedDelDate].Value = po.Dates.RevisedScheduledDeliveryDate;
                    }
                    
                    // Vendor Class
                    ws.Cells[nextRow, dCol.VendorAccount].Value2 = po.Vendor.Code;
                    ws.Cells[nextRow, dCol.VendorName].Value2 = po.Vendor.Name;

                    nextRow++;
                    M.updateMetrics.QTotalUpdatedLines++;
                }
            }

            if (M.Dates.IsDatesToUpdateInExpediteReport)
            {
                M.kaxlApp.ErrorTracker.ProgramStage = "Updating Dates in Expedite Report";

                M.updateMetrics.QUpdatedReceivedDates = M.Dates.QReceivedDatesToUpdate;
                M.updateMetrics.QUpdatedRevisedDeliveryDates = M.Dates.QRevisedScheduledDeliveryDatesToUpdate;
                M.Dates.UpdateDatesOnExpediteReport();
            }

            WS expRep = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.MasterData];

            //Update vendor list with vendor numbers that aren't in dictionary
            if (M.VendorDict.IsVendorNumbersThatArentInDict())
            {
                M.kaxlApp.ErrorTracker.ProgramStage = "Updating vendor names in vendor list";

                int col = (int)Master.MasterDataColumnsE.VendorAccount;
                int NR = KAXL.LastRow(expRep, col) + 1;

                List<string> vendorNamesNotInDictionary = M.VendorDict.VendorNumbersThatArentInDictL();

                foreach (string VendorNumber in vendorNamesNotInDictionary)
                {
                    expRep.Cells[NR, col].Value2 = VendorNumber;
                    NR++;
                }
            }
            // Update Item List with item numbers not in dictionary
            if (M.ItemDict.IsItemsThatArentInDict())
            {
                M.kaxlApp.ErrorTracker.ProgramStage = "Updating Item's that aren't in dictionary";

                int col = (int)Master.MasterDataColumnsE.ItemNum;
                int NR = KAXL.LastRow(expRep, col) + 1;

                List<string> itemNumbersNotInDictionary = M.ItemDict.GetItemNumbersThatArentInDictList();

                foreach (string item in itemNumbersNotInDictionary)
                {
                    if(item != null)
                    {
                        expRep.Cells[NR, col].Value2 = item;
                        NR++;
                    }                    
                }
            }
            M.stopWatch.EndTime = DateTime.Now;
        }        
    }
}
