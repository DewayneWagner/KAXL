using DKAExcelStuff;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SN = EXPREP_V2.Master.SheetNamesE;
using WS = Microsoft.Office.Interop.Excel.Worksheet;

namespace EXPREP_V2
{
    public class ScrubbedPOLine
    {
        public ScrubbedPOLine() { }

        Master m;
        private List<ScrubbedPOLine> _scrubbedPOLine;

        public ScrubbedPOLine(Master master)
        {
            m = master;
            _scrubbedPOLine = new List<ScrubbedPOLine>();
            LoadScrubbedPOLineList();
        }

        public ScrubbedPOLine this[int i]
        {
            get => _scrubbedPOLine.Count >= i ? _scrubbedPOLine[i] : null;
            set => _scrubbedPOLine[i] = value;
        }

        public List<ScrubbedPOLine> GetList() => _scrubbedPOLine;

        public string PONum { get; set; }
        public double LineNumber { get; set; }
        public double Quantity { get; set; }
        public Cash Cash { get; set; }
        public Vendor Vendor { get; set; }
        public Category Category { get; set; }
        public Item ItemX { get; set; }
        public Source Source { get; set; }
        public AllDates Dates { get; set; }

        public string WH { get; set; }
        public string GetWH(string wh) => wh.Length >= 3 && wh != null ? wh.Substring(0, 3) : null;
            
        public Status Status { get; set; }
        public String Direct { get; set; }
        public string Entity => PONum.Substring(0, 4);
        public bool ICO { get; set; }
        public bool IsLineInExpRep => (m.PODictionaryInExpRep.IsDuplicate(PONum, Math.Floor(LineNumber))) ? true : false;
        public bool IsReceived { get; set; }

        public string Receiver
        {
            get
            {
                switch (WH)
                {
                    case ("MIS"):
                        return "Arbie";
                    case ("COL"):
                        return "Michael";
                    case ("DAL"):
                        return "Charles";
                    case ("CAL"):
                        return "Dave";
                }
                return null;
            }
        }

        private bool expediteRequired;
        public bool ExpediteRequired
        {
            get
            {
                try
                {
                    if (Category.CleanCategory == "Subcontractor" ||
                    Vendor.Name == "McMaster-Carr Supply Co." ||
                    Vendor.Name == "McMaster-Carr" ||
                    Vendor.Name == "Mouser Electronics, Inc" ||
                    Vendor.Name == "Mouser Electronics" ||
                    Source.CreatedBy == "DarrenM" ||
                    ICO)
                    {
                        expediteRequired = false;
                        return expediteRequired;
                    }
                    else
                    {
                        expediteRequired = true;
                        return expediteRequired;
                    }
                }
                catch
                {
                    expediteRequired = true;
                    return expediteRequired;
                }
            }
            set => expediteRequired = value;
        }

        private void LoadScrubbedPOLineList()
        {
            for (int sheet = (int)SN.PTCA; sheet <= (int)SN.HMCA; sheet++)
            {
                m.kaxlApp.WS = m.kaxlApp.WB.Sheets[sheet];

                WS ws = m.kaxlApp.WS;
                var k = m.kaxlApp.KAXL_RG;
                k = new KAXLApp.KAXLRange(m.kaxlApp, RangeType.WorkSheet);
                SourceColID sColID = new SourceColID(ws);
                m.kaxlApp.ErrorTracker.ProgramStage = "Reading " + Convert.ToString((SN)sheet);
                
                for (int r = 1; r < k.Row.End; r++)
                {
                    m.kaxlApp.ErrorTracker.Row = r;
                    string poNumber, key;
                    double lineNumber;

                    try
                    {
                        poNumber = (string)k[r, sColID.PurchaseOrder];
                        lineNumber = Convert.ToDouble(k[r, sColID.LineNumber]);
                        key = poNumber + Convert.ToString(Math.Round(lineNumber,0));

                        AllDates dates = new AllDates()
                        {
                            OriginalScheduledDelivery = KAXL.ReadDateTime(k[r, sColID.OrigSchedDelDate]),
                            POCreated = KAXL.ReadDateTime(k[r, sColID.CreatedDate]),
                            RevisedScheduledDeliveryDate = KAXL.ReadDateTime(k[r, sColID.RevisedSchedDelDate]),
                        };

                        Status status = new Status((string)k[r, sColID.LineStatus], m, poNumber, Convert.ToString(lineNumber), (string)k[r, sColID.ApprovalStatus]);

                        if(m.PODictionaryInExpRep.IsDuplicate(key))
                        {
                            CheckAndUpdateReceivedAndRevisedDate(m, (m.PODictionaryInExpRep[key].ExpRepXLLineNum), dates, key, status.CleanStatus);
                        }
                        else if (status.CleanStatus != Status.CleanStatusE.Canceled && status.CleanStatus != Status.CleanStatusE.Draft)
                        {
                            Item itemX = m.ItemDict[Convert.ToString(k[r, sColID.ItemNumber])];
                            string procurementCategory = (string)k[r, sColID.ProcurementCategory];
                            Category category = new Category(procurementCategory, itemX, m);
                            Source source = new Source((string)k[r, sColID.AttentionInformation]);
                            double quantity = Convert.ToDouble(k[r, sColID.Quantity]);
                            Cash cash = new Cash((string)k[r, sColID.Currency], Convert.ToDouble(k[r, sColID.NetAmount]), dates.POCreated, m, quantity);
                            Vendor vendor = m.VendorDict[(string)k[r, sColID.VendorAccount]];
                            string wh = (string)k[r, sColID.Warehouse];
                            string direct = itemX.Desc == null ? "Indirect" : "Direct";
                            bool ico = vendor.Name != null && (vendor.Code.Length == 4) ? true : false;

                            _scrubbedPOLine.Add(new ScrubbedPOLine()
                            {
                                LineNumber = lineNumber,
                                Cash = cash ?? new Cash(),
                                ItemX = itemX ?? new Item(),
                                Category = category ?? new Category(),
                                Dates = dates ?? new AllDates(),
                                PONum = poNumber,
                                Quantity = quantity,
                                Source = source ?? new Source(),
                                Status = status ?? new Status(),
                                Vendor = vendor ?? new Vendor(),
                                WH = wh != null ? GetWH(wh) : "No WH",
                                Direct = direct,
                                ICO = ico,
                            });

                            var p = _scrubbedPOLine[_scrubbedPOLine.Count - 1];
                            if (p.Source.IsMultiLinePO)
                            {
                                p.LineNumber = lineNumber + 0.1;
                                int q = p.Source.QSourcesInList;

                                for (int i = 1; i < q; i++)
                                {
                                    p.LineNumber += 0.1;
                                    p.Quantity = 0;
                                    p.Cash.NetAmount = 0;
                                    p.Cash = Cash.ZeroedOutCash();
                                    p.Source = p.Source[i];

                                    _scrubbedPOLine.Add(p);
                                }
                            }
                        }
                    }
                    catch
                    {
                        m.kaxlApp.ErrorTracker.AddNewError(Convert.ToString((SN)sheet) + ", Row" + Convert.ToString(r));
                    }                    
                }
            }
        }
        private int Last() => _scrubbedPOLine.Count - 1;
        
        private void CheckAndUpdateReceivedAndRevisedDate(Master m, int row, AllDates dates, string key, Status.CleanStatusE status)
        {
            PODictionaryInExpRep po = m.PODictionaryInExpRep[key];

            if (po.MostRecentRevisedDeliveryDate != dates.RevisedScheduledDeliveryDate && dates.RevisedScheduledDeliveryDate != DateTime.MinValue)
            {
                m.Dates.AddDateToExpRepUpdateList(row, dates.RevisedScheduledDeliveryDate);
            }
            if (!po.IsReceivedDatePresent && status == Status.CleanStatusE.Received)
            {
                m.Dates.AddDateToExpRepUpdateList(row, DateTime.MinValue, true);
            }
        }
        public int ListQ => _scrubbedPOLine.Count;
    }
}
