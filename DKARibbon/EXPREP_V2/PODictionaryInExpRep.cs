using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using DKAExcelStuff;

namespace EXPREP_V2
{
    public class PODictionaryInExpRep
    {
        public PODictionaryInExpRep() { }

        Master m;
        private readonly Dictionary<string, PODictionaryInExpRep> _poDictionaryInExpRep;

        public PODictionaryInExpRep(Master master)
        {
            m = master;
            _poDictionaryInExpRep = new Dictionary<string, PODictionaryInExpRep>();
            LoadDictionaryWithPOs();
        }

        public string PONum { get; set; }
        public double POLineNum { get; set; }
        public Status Status { get; set; }
        public int ExpRepXLLineNum { get; set; }
        public DateTime MostRecentRevisedDeliveryDate { get; set; }
        public bool IsReceivedDatePresent { get; set; }

        private void LoadDictionaryWithPOs()
        {
            m.kaxlApp.WS = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];

            WS ws = m.kaxlApp.WS;
            var k = m.kaxlApp.KAXL_RG;
            k = new KAXLApp.KAXLRange(m.kaxlApp, RangeType.WorkSheet);

            m.kaxlApp.ErrorTracker.ProgramStage = "Reading Expedite Report";

            string poNum, key;
            double lineNum;

            try
            {
                for (int r = KAXL.FindFirstRowAfterHeader(ws); r < k.Row.End; r++)
                {
                    m.kaxlApp.ErrorTracker.Row = r;

                    poNum = (string)k[r, m.ExpRepColumn.PONumber];
                    lineNum = Math.Round((double)k[r, m.ExpRepColumn.LineNumber]);
                    key = poNum + Convert.ToString(lineNum);
                    DateTime receivedDate = KAXL.ReadDateTime(k[r, m.ExpRepColumn.RecDate]);
                    IsReceivedDatePresent = receivedDate == DateTime.MinValue ? false : true;

                    if (_poDictionaryInExpRep.ContainsKey(key))
                    {
                        key = null;
                    }
                    else
                    {
                        DateTime revDate = KAXL.ReadDateTime(k[r, m.ExpRepColumn.RevisedSchedDelDate]);
                        bool isReceivedDatePresent = (k[r, m.ExpRepColumn.RecDate] is null) ? false : true;

                        Status status = new Status()
                        {
                            ExpRepStatus = (string)k[r, m.ExpRepColumn.Status],
                            PONum = poNum,
                        };

                        _poDictionaryInExpRep.Add(key, new PODictionaryInExpRep()
                        {
                            ExpRepXLLineNum = r,
                            IsReceivedDatePresent = isReceivedDatePresent,
                            MostRecentRevisedDeliveryDate = revDate,
                            POLineNum = lineNum,
                            PONum = poNum,
                            Status = status
                        });

                        string itemNum = Convert.ToString(k[r, m.ExpRepColumn.ItemNumber]);
                        string vendName = (string)k[r, m.ExpRepColumn.VendorName];

                        if (itemNum != null && (string)k[r, m.ExpRepColumn.ItemDescription] == null)
                        {
                            // do this if there is an item number, but not description
                            Item i = m.ItemDict[itemNum];
                            ws.Cells[r, m.ExpRepColumn.ItemDescription].Value2 = i.Desc;
                        }

                        // check if vendor name is present, and add to update list if not.
                        if (vendName == null)
                        {
                            ws.Cells[r, m.ExpRepColumn.VendorName].Value2 = m.VendorDict[vendName];
                        }
                    }
                }
            }
            catch
            {
                m.kaxlApp.ErrorTracker.AddNewError("Reading ExpRep, Line #" + Convert.ToString(m.kaxlApp.ErrorTracker.Row));
            }
        }
        public bool IsDuplicate(string poNum, double lineNum) =>
            (_poDictionaryInExpRep.ContainsKey(poNum + Convert.ToString(Math.Floor(lineNum)))) ? true : false;

        public bool ContainsKey(string key) => _poDictionaryInExpRep.ContainsKey(key);

        public PODictionaryInExpRep this[string key]
        {
            get => key != null && _poDictionaryInExpRep.ContainsKey(key) ? _poDictionaryInExpRep[key] : null;
            set => _poDictionaryInExpRep[key] = value;
        }
    }
}
