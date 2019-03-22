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
        public DateTime MostRecRevDate { get; set; }
        public bool IsReceivedDatePresent { get; set; }

        private void LoadDictionaryWithPOs()
        {
            m.kaxlApp.WS = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];

            WS ws = m.kaxlApp.WS;
            var k = m.kaxlApp.KAXL_RG;
            k = new KAXLApp.KAXLRange(m.kaxlApp, RangeType.WorkSheet);

            m.errorTracker.Process = "Reading Expedite Report";
            string poNum, key;
            double lineNum;

            for (int r = KAXL.FindFirstRowAfterHeader(ws); r < k.Row.End; r++)
            {
                m.errorTracker.LineNumber = Convert.ToString(r);
                poNum = (string)k[r, m.ExpRepColumn.PONumber];
                lineNum = Math.Round((double)k[r, m.ExpRepColumn.LineNumber]);
                key = poNum + Convert.ToString(lineNum);
                DateTime mostRecentRevisedDeliveryDate;

                if (_poDictionaryInExpRep.ContainsKey(key))
                {
                    key = null;
                }
                else
                {
                    var rev = k[r, m.ExpRepColumn.RevisedSchedDelDate];

                    if (rev != null)
                        mostRecentRevisedDeliveryDate = Convert.ToDateTime(rev);
                    else
                        mostRecentRevisedDeliveryDate = DateTime.MinValue;

                    //var mostRecRevDate = k[r, m.ExpRepColumn.RevisedSchedDelDate] == null ? null : (DateTime)k[r, m.ExpRepColumn.RevisedSchedDelDate];
                    _poDictionaryInExpRep.Add(key, new PODictionaryInExpRep()
                    {
                        ExpRepXLLineNum = r,
                        IsReceivedDatePresent = (k[r,m.ExpRepColumn.RecDate] is null) ? false : true,
                        MostRecRevDate = mostRecentRevisedDeliveryDate,
                        POLineNum = lineNum,
                        PONum = poNum,
                        Status = new Status()
                        {
                            ExpRepStatus = (string)k[r,m.ExpRepColumn.Status],
                            PONum = poNum,
                        },
                    });

                    string itemNum = Convert.ToString(k[r, m.ExpRepColumn.ItemNumber]);
                    string vendName = (string)k[r, m.ExpRepColumn.VendorName];

                    if (itemNum != null && (string)k[r,m.ExpRepColumn.ItemDescription] == null)
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
