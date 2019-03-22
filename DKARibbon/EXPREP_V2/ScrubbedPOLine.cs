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

        public string PONum { get; set; }
        public double LineNumber { get; set; }
        public double Quantity { get; set; }
        public Cash Cash { get; set; }
        public Vendor Vendor { get; set; }
        public Category Category { get; set; }
        public Item ItemX { get; set; }
        public Source Source { get; set; }
        public Dates Dates { get; set; }
        public string WH => WH?.Substring(0, 3);
        public Status Status { get; set; }
        public String Direct => ItemX.Num == null ? "Indirect" : "Direct";
        public string Entity => PONum.Substring(0, 4);
        public bool ICO => Vendor != null && (Vendor.Code.Length == 4) ? true : false;
        public bool IsLineInExpRep => (m.PODictionaryInExpRep.IsDuplicate(PONum, Math.Floor(LineNumber))) ? true : false;

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
                m.errorTracker.Process = "Reading " + Convert.ToString((SN)sheet);

                for (int r = KAXL.FindFirstRowAfterHeader(ws); r < k.Row.End; r++)
                {
                    string poNumber = (string)k[r, sColID.PurchaseOrder];
                    double lineNumber = (double)k[r, sColID.LineNumber];
                    string key = poNumber + Convert.ToString(lineNumber);

                    if (m.PODictionaryInExpRep.ContainsKey(key))
                    {
                        CheckAndUpdateReceivedAndRevisedDate();
                    }
                    else
                    {
                        Item itemX = m.ItemDict[Convert.ToString(k[r, sColID.ItemNumber])];
                        string procurementCategory = (string)k[r, sColID.ProcurementCategory];
                        Category category = new Category(procurementCategory, itemX, m);
                        RevisedSchedDeliveryDate revSchedDelTime = new RevisedSchedDeliveryDate(ScrubDate(k[r, sColID.RevisedSchedDelDate]), m, poNumber, lineNumber, r);

                        _scrubbedPOLine.Add(new ScrubbedPOLine()
                        {
                            LineNumber = lineNumber,
                            Cash = new Cash()
                            {
                                Currency = (string)k[r, sColID.Currency],
                                NetAmount = (double)k[r, sColID.NetAmount]
                            },
                            ItemX = itemX,
                            Category = category,
                            Dates = new Dates()
                            {
                                OrigSchedDelDate = ScrubDate(k[r, sColID.OrigSchedDelDate]),
                                POCreatedDate = ScrubDate(k[r, sColID.CreatedDate]),
                                RevisedSchedDelDate = revSchedDelTime,
                            },                            
                            PONum = poNumber,
                            Quantity = (double)k[r, sColID.Quantity],
                            Source = new Source((string)k[r, m.SColID.AttentionInformation]),
                            Status = new Status((string)k[r, sColID.LineStatus], m, poNumber, Convert.ToString(lineNumber), (string)k[r, sColID.ApprovalStatus]),
                            Vendor = new Vendor() { Code = (string)k[r, sColID.VendorAccount] },
                        });
                        var p = _scrubbedPOLine[Last()];
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
            }
        }
        private int Last() => _scrubbedPOLine.Count;

        private DateTime ScrubDate(object dt)
        {
            if (dt != null)
                return Convert.ToDateTime(dt);
            else
                return DateTime.MinValue;
        }

        private void CheckAndUpdateReceivedAndRevisedDate()
        {
            PODictionaryInExpRep po = m.PODictionaryInExpRep[PONum + Math.Floor(LineNumber)];
            DateTime revisedSchedDelDate = Dates.RevisedSchedDelDate.MostRecentShedDeliveryDate;

            if (po.MostRecRevDate != revisedSchedDelDate || revisedSchedDelDate == DateTime.MinValue)
            {
                m.RevisedSchedDelDatesToUpdate.AddToUpdateList(po.ExpRepXLLineNum, po.MostRecRevDate);
            }
            if (!po.IsReceivedDatePresent && Status.CleanStatus == "Received")
            {
                m.ReceivedDateList.AddToUpdateList(po.ExpRepXLLineNum);
            }
        }

        public int ListQ => _scrubbedPOLine.Count;
    }
}
