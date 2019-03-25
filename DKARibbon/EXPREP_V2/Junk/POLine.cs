using System;
using System.Collections.Generic;
using System.Collections;
using DKAExcelStuff;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;

namespace EXPREP_V2
{
    //public class POLine
    //{
    //    public Master M;
    //    public POLine(Master m) => M = m;

    //    public POLine() { }

    //    // constructor to use for creating po lines from source data
    //    public POLine(List<string> poLine, Master m) 
    //    {
    //        M = m;
    //        var c = M.SColID;

    //        // first load PO Number and PO Line
    //        PONum = poLine[c.PurchaseOrder];
    //        LineNumber = Convert.ToDouble(poLine[c.LineNumber]);

    //        Dates = new Dates(poLine[c.RevisedSchedDelDate], poLine[c.OrigSchedDelDate], 
    //            poLine[c.CreatedDate]);
    //        Status = new Status(poLine[c.LineStatus],M,PONum,Convert.ToString(LineNumber),poLine[c.ApprovalStatus]);

    //        if (IsLineInExpRep)
    //        {
    //            CheckAndUpdateReceivedAndRevisedDate();
    //        }            
    //        else
    //        {
    //            Source = new Source(poLine[c.AttentionInformation]);
    //            Cash = new Cash(poLine[c.Currency], poLine[c.UnitPrice], poLine[c.NetAmount], poLine[c.CreatedDate], M);
    //            Vendor = M.VendorDict[poLine[c.VendorAccount]];                
    //            Item = M.ItemDict[poLine[c.ItemNumber]];
    //            Category = new Category(poLine[c.ProcurementCategory], Item, M);
    //            Quantity = Convert.ToDouble(poLine[c.Quantity]);                

    //            WH = FormatWH(poLine[c.Warehouse]);
    //            Receiver = FormatReceiver();
    //        }            
    //    }
    //    private void CheckAndUpdateReceivedAndRevisedDate() 
    //    {
    //        PODictionaryInExpRep po = M.PODictionaryInExpRep[PONum + Math.Floor(LineNumber)];
    //        DateTime revisedSchedDelDate = Dates.RevisedSchedDelDate.MostRecentShedDeliveryDate;

    //        if (po.MostRecRevDate != revisedSchedDelDate || revisedSchedDelDate == DateTime.MinValue)
    //        {
    //            M.RevisedSchedDelDatesToUpdate.AddToUpdateList(po.ExpRepXLLineNum, po.MostRecRevDate);
    //        }                
    //        if (!po.IsReceivedDatePresent && Status.CleanStatus == "Received")
    //        {
    //            M.ReceivedDateList.AddToUpdateList(po.ExpRepXLLineNum);
    //        }            
    //    }

    //    // this will be the class for each line of the expedite report
    //    public string PONum { get; set; }
    //    public double LineNumber { get; set; } 
    //    public double Quantity { get; set; }                
    //    public Cash Cash { get; set; }
    //    public Vendor Vendor { get; set; }
    //    public Category Category { get; set; }
    //    public Item Item { get; set; }
    //    public Source Source { get; set; }
    //    public Dates Dates { get; set; }
    //    public string WH { get; set; }
    //    public string Receiver { get; set; }
    //    public Status Status { get; set; }
    //    public String Direct => Item.Num == null ? "Indirect" : "Direct";

    //    public string Entity 
    //    {
    //        get => entity = PONum != null ? PONum.Substring(0, 4) : null;
    //        set => entity = value;
    //    }

    //    public bool ICO => Vendor != null && (Vendor.Code.Length == 4) ? true : false;

    //    // when PODictionaryInExpRep is created - this tracks which row on the spreadsheet a certain record is
    //    public int ExpRepXLLineNum { get; set; }

    //    //----------------- METHOD SECTION ----------------------------------------------
    //    private string FormatWH(string wh) => wh?.Substring(0, 3);

    //    private string FormatReceiver()
    //    {
    //        switch (WH)
    //        {
    //            case ("MIS"):
    //                return "Arbie";
    //            case ("COL"):
    //                return "Michael";
    //            case ("DAL"):
    //                return "Charles";
    //            case ("CAL"):
    //                return "Dave";
    //        }
    //        return null;
    //    }

    //    private bool expediteRequired;
    //    public bool ExpediteRequired
    //    {            
    //        get
    //        {
    //            try
    //            {
    //                if (Category.CleanCategory == "Subcontractor" ||
    //                Vendor.Name == "McMaster-Carr Supply Co." ||
    //                Vendor.Name == "McMaster-Carr" ||
    //                Vendor.Name == "Mouser Electronics, Inc" ||
    //                Vendor.Name == "Mouser Electronics" ||
    //                //Vendor.Name == "Digi-Key Corporation" ||
    //                //Vendor.Name == "Digi-Key Corp." ||
    //                //Vendor.Name == "Newark" ||
    //                //Vendor.Name == "Uline Shipping Supplies" ||
    //                //Vendor.Name == "Uline Canada Corporation" ||
    //                //Vendor.Name == "ULINE" ||
    //                Source.CreatedBy == "DarrenM" ||
    //                ICO)
    //                {
    //                    expediteRequired = false;
    //                    return expediteRequired;
    //                }
    //                else
    //                {
    //                    expediteRequired = true;
    //                    return expediteRequired;
    //                }
    //            }
    //            catch
    //            {
    //                expediteRequired = true;
    //                return expediteRequired;
    //            }
    //        }
    //        set => expediteRequired = value;
    //    }
    //    public bool IsLineInExpRep => (M.PODictionaryInExpRep.IsDuplicate(PONum, Math.Floor(LineNumber))) ? true : false;         
    //}
    //public class POLinesList : IEnumerator, IEnumerable
    //{
    //    private List<POLine> poLinesL;
    //    int position = -1;
    //    private Master M;

    //    public POLinesList(Master m)
    //    {
    //        poLinesL = new List<POLine>();
    //        M = m;
    //        ReadPOLineData();   
    //    }

    //    public POLine this[int i]
    //    {
    //        get => poLinesL.Count >= i ? poLinesL[i] : null;  
    //        set => poLinesL[i] = value;
    //    }
    //    public void AddPOLine(POLine poLine) => poLinesL.Add(poLine);
    //    public int POLineListQ => poLinesL.Count;

    //    public IEnumerator GetEnumerator() => (IEnumerator)this;

    //    public bool MoveNext()
    //    {
    //        position++;
    //        return (position < poLinesL.Count);
    //    }
    //    public void Reset() => position = 0;
    //    public object Current => poLinesL[position];

    //    public void ReadPOLineData()
    //    {
    //        int startRow, LR, LC;
    //        for (int i = (int)Master.SheetNamesE.PTCA; i <= (int)Master.SheetNamesE.HMCA; i++)
    //        {
    //            WS ws = M.kaxlApp.WB.Sheets[Convert.ToString((Master.SheetNamesE)i)];
    //            SourceColID c = new SourceColID(ws);

    //            M.SColID = c;
    //            startRow = 2;

    //            LR = KAXL.LastRow(ws, 1);
    //            LC = KAXL.LastCol(ws, 1);
                
    //            List<string> rowDataL = new List<string>() { null };
    //            int iRow;
                
    //            for (iRow = startRow; iRow <= LR; iRow++)
    //            {
    //                M.errorTracker.Process = "Reading " + (Master.SheetNamesE)i;
    //                M.errorTracker.LineNumber = Convert.ToString(iRow);

    //                rowDataL.Clear();
    //                rowDataL.Add("null");

    //                // for identifying what row the program is erroring-out on
    //                //ws.Cells[1, 1].Value2 = iRow;

    //                for (int iCol = 1; iCol <= LC; iCol++)
    //                {
    //                    var val = ws.Cells[iRow, iCol].Value;

    //                    if (val is string)
    //                        rowDataL.Add(val);
    //                    else
    //                        rowDataL.Add(Convert.ToString(val));
    //                }
    //                POLine po = new POLine(rowDataL, M);

    //                if (!po.IsLineInExpRep)
    //                {
    //                    if (po.Source.IsMultiLinePO)
    //                    {
    //                        int q = po.Source.QSourcesInList;
    //                        Source[] tempArray = new Source[q];
    //                        for (int iii = 0; iii < q; iii++)
    //                        {
    //                            tempArray[iii] = po.Source[iii];
    //                        }
    //                        double poLineNum = po.LineNumber;
    //                        for (int ii = 0; ii < q; ii++)
    //                        {
    //                            if (ii == 0)
    //                            {
    //                                po.LineNumber += 0.1;
    //                                po.Source = tempArray[ii];
    //                                poLinesL.Add(po);
    //                            }
    //                            else
    //                            {
    //                                Source pos = tempArray[ii];
    //                                AddMultiLinePOToPOList(po, pos, q);
    //                            }
    //                        }
    //                    }
    //                    else
    //                    {
    //                        poLinesL.Add(po);
    //                    }
    //                }                   
    //            }
    //            ws.Cells.ClearContents();
    //        }
    //    }
    //    private void AddMultiLinePOToPOList(POLine po, Source pos, int q)
    //    {            
    //        double lineNumber = po.LineNumber;
    //        double quantity = po.Quantity;
    //        double lineNum = po.LineNumber;

    //        for (int i = 1; i < q; i++)
    //        {                
    //            lineNum += 0.1;

    //            POLine newPO = new POLine();

    //            Cash cash = new Cash(true);
    //            newPO.Cash = cash;

    //            // properties to match PO
    //            newPO.Category = po.Category;
    //            newPO.Dates = po.Dates;
    //            newPO.Entity = po.Entity;
    //            newPO.ExpediteRequired = po.ExpediteRequired;
    //            newPO.Item = po.Item;
    //            newPO.PONum = po.PONum;
    //            newPO.WH = po.WH;
    //            newPO.Status = po.Status;
    //            newPO.Vendor = po.Vendor;
    //            newPO.Receiver = po.Receiver;
    //            newPO.Cash.Currency = po.Cash.Currency;

    //            // for multiline PO's
    //            newPO.Quantity = 0;
    //            newPO.Source = pos;
    //            newPO.LineNumber = lineNum;

    //            poLinesL.Add(newPO);
    //        }
    //    }
    //}
}
