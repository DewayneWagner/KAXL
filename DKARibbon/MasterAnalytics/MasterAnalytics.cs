using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using Microsoft.Office.Tools.Ribbon;
using NR = Microsoft.Office.Tools.Excel.NamedRange;
using DKARibbon;
using SD = System.Data;
using System.Reflection;
using System.Windows.Forms;

namespace DKAExcelStuff
{ }
//    public class MasterAnalytics
//    {
//        private enum ItemDescriptionHeadings {Nada,ItemNum,ItemDesc,Tot}
//        private enum ExRateSourceHeadings { Nada, Currency, Year, Month, ExRate, Tot }

//        public static List<ExRate> ExRatesL = new List<ExRate>();
//        public static List<Vendor> VendorNamesL = new List<Vendor>();
//        public static Dictionary<string, ItemDesc> itemDict = new Dictionary<string, ItemDesc>();

//        public static string ExpPath = null;
//        public static string MasterDataPath = null;

//        public static void MasterAnalyticsStart(Worksheet ws)
//        {
//            DateTime StartTime = DateTime.Now;

//            int LR = KAXL.LastRow(ws,1) + 1;
//            int LC = KAXL.LastCol(ws,1);

//            /****************************************************************************************
//             * Determine which column of source data is applicable to each field
//             ***************************************************************************************/

//            SourceColID sCol = new SourceColID(ws);

//            /********************************************************************************************/
            
//            List<string> sourceItemNumsL = new List<string>();
            
//            for (int i = 2; i <= LR; i++)
//            {
//                var itemNum = ws.Cells[i, sCol.ItemNumber].Value2;
//                if (itemNum is string && itemNum != null && !sourceItemNumsL.Contains(itemNum))
//                    sourceItemNumsL.Add(itemNum);
//                else if (itemNum is double && !sourceItemNumsL.Contains(Convert.ToString(itemNum)))
//                    sourceItemNumsL.Add(Convert.ToString(itemNum));
//            }

//            ReferenceData refData = new ReferenceData(sourceItemNumsL);

//            ExRatesL = refData.ExRateL;
//            VendorNamesL = refData.VendorL;

//            itemDict = refData.ItemDict;
            
//            /*******************************************************************************************/

//            // load dirty array with values
//            string[,] dirtyArray = KAXL.LoadDirtyArr(ws, (LR+1), (LC+1));

//            // apply standard fields to each transaction in dirty array
//            List<Transaction> TransactionObjectsL = new List<Transaction>() { null };

//            for (int r = 2; r <= LR; r++)
//            {
//                Transaction trans = new Transaction();
//                trans.Status = dirtyArray[r, sCol.LineStatus];

//                if (trans.Status != "Cancelled")
//                    TransactionObjectsL.Add(trans);
//                else
//                    continue;

//                trans.VendorAccount = dirtyArray[r, sCol.VendorAccount];
//                trans.PO = dirtyArray[r, sCol.PurchaseOrder];
//                trans.ProcurementCategory = dirtyArray[r, sCol.ProcurementCategory];
//                trans.Site = dirtyArray[r, sCol.Site];
//                trans.Warehouse = dirtyArray[r, sCol.Warehouse];
//                trans.Currency = dirtyArray[r, sCol.Currency];
//                trans.LineNumber = Convert.ToInt32(dirtyArray[r,sCol.LineNumber]);
//                trans.ItemNumber = dirtyArray[r, sCol.ItemNumber];
//                trans.Quantity = Convert.ToDecimal(dirtyArray[r,sCol.Quantity]);
//                trans.UnitPrice = Convert.ToDecimal(dirtyArray[r,sCol.UnitPrice]);
//                trans.NetAmount = Convert.ToDouble(dirtyArray[r,sCol.NetAmount]);
                
//                trans.AttentionInformation = dirtyArray[r, sCol.AttentionInformation];

//                trans.POCreatedDate = KAXL.ReadDateTime(dirtyArray[r,sCol.CreatedDate]);
//                trans.OriginalScheduledDeliveryDate = KAXL.ReadDateTime(dirtyArray[r,sCol.ConfirmedDate]);
//                trans.RevisedSheduledDeliveryDate = KAXL.ReadDateTime(dirtyArray[r,sCol.DeliveryDate]);

//                if (trans.AttentionInformation != null)
//                {
//                    try
//                    {
//                        POSource poSource = new POSource(trans.AttentionInformation);

//                        for (int i = 0; i < poSource.QPOSources; i++)
//                        {
//                            if (i == 0)
//                            {
//                                trans.POSourceType = poSource.POSourceA[(int)POSource.POSourceItemsE.POSourceType, i];
//                                trans.POSourceCode = poSource.POSourceA[(int)POSource.POSourceItemsE.POSourceCode, i];
//                                trans.CreatedBy = poSource.POSourceA[(int)POSource.POSourceItemsE.CreatedBy, i];
//                                trans.Requester = poSource.POSourceA[(int)POSource.POSourceItemsE.Requester, i];
//                            }
//                            else
//                            {
//                                Transaction multiLineTrans = new Transaction();
//                                TransactionObjectsL.Add(multiLineTrans);

//                                multiLineTrans.VendorAccount = trans.VendorAccount;
//                                multiLineTrans.PO = trans.PO;
//                                multiLineTrans.LineNumber = trans.LineNumber;
//                                multiLineTrans.ItemNumber = trans.ItemNumber;
//                                multiLineTrans.AttentionInformation = trans.AttentionInformation;
//                                multiLineTrans.POCreatedDate = trans.POCreatedDate;
//                                multiLineTrans.OriginalScheduledDeliveryDate = trans.OriginalScheduledDeliveryDate;
//                                multiLineTrans.Currency = trans.Currency;
//                                multiLineTrans.POCreatedDate = trans.POCreatedDate;
//                                multiLineTrans.ProcurementCategory = trans.ProcurementCategory;
//                                multiLineTrans.Quantity = trans.Quantity;
//                                multiLineTrans.RevisedSheduledDeliveryDate = trans.RevisedSheduledDeliveryDate;
//                                multiLineTrans.Site = trans.Site;
//                                multiLineTrans.Status = trans.Status;
//                                multiLineTrans.UnitPrice = trans.UnitPrice;
//                                multiLineTrans.Warehouse = trans.Warehouse;

//                                multiLineTrans.POSourceCode = poSource.POSourceA[(int)POSource.POSourceItemsE.POSourceCode, i];
//                                multiLineTrans.Requester = poSource.POSourceA[(int)POSource.POSourceItemsE.Requester, i];
//                                multiLineTrans.Quantity = 0;
//                                multiLineTrans.NetAmount = 0;
//                            }
//                        }                    
//                    }
//                    catch
//                    {
//                        trans.POSourceType = "Error";
//                        trans.POSourceCode = "Error";
//                        trans.CreatedBy = "Error";
//                        trans.Requester = "Error";
//                    }
//                }
//            }

//            int rowQ = TransactionObjectsL.Count;
//            AddToExpRep.AddTransListToExpRep(TransactionObjectsL);

//            DateTime endTime = DateTime.Now;
//            TimeSpan duration = endTime - StartTime;

//            MessageBox.Show("Expedite Report Finished Updating, JackAss." + "\n\n" +
//                "Time to Complete:  " + duration + "\n\n" +
//                "Number of Lines Updated:  " + rowQ);           
            
//            // Code that works!!!!!!!!!!
//            //SD.DataTable dt = CreateDataTable.ToDataTable(TransactionObjectsL);                      
//            //DKAWrite dkawrite = new DKAWrite();
//            //dkawrite.WriteDataFromTable(dt, ws);
//            //xl.FormatAsTable("POData"); 
//        }        
//    }
//}

