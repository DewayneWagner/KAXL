using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using DKAExcelStuff;
using Microsoft.Office.Interop.Excel;

namespace EXPREP_V2
{
    public class PODictionaryInExpRep
    {        
        private enum RequiredFields { PoNumberCol, LineNumberCol, RecDateCol, RevDateCol, StatusCol, Total}

        public PODictionaryInExpRep() { }

        Master m;
        private readonly Dictionary<string, PODictionaryInExpRep> _poDictionaryInExpRep;

        public PODictionaryInExpRep(Master master)
        {
            m = master;
            _poDictionaryInExpRep = new Dictionary<string, PODictionaryInExpRep>();

            //LoadDictionaryWithPOs();
            LoadDictionary();
        }

        public string PONum { get; set; }

        private double _poLineNum;
        public double POLineNum
        {
            get => _poLineNum;
            set => _poLineNum = Math.Round((double)value, 0);
        }

        private string _key;
        public string Key
        {
            get => _key;
            set => _key = PONum + Convert.ToString(POLineNum);
        }

        public Status Status { get; set; }
        public int ExpRepXLLineNum { get; set; }
        public DateTime MostRecentRevisedDeliveryDate { get; set; }
        public bool IsReceivedDatePresent { get; set; }

        // array of array method
        private void LoadDictionary()
        {
            WS ws = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            int firstRow = 3;
            int lastRow = KAXL.LastRow(ws,1);
            int qRows = lastRow - firstRow + 1;
            int dummyColumnIndex = 1;

            m.kaxlApp.ErrorTracker.ProgramStage = "Reading Expedite Report";

            List<int> _colNumsOfRequiredFields = LoadListOfColNumsOfReqFields();

            // ends-up being a 0-based indexed array
            object[,] _objectArray = new object[qRows, (int)RequiredFields.Total];

            // is read as a 1-based indexed array
            object[,] _2DArrayOf1DData;

            for (int i = 0; i < (int)RequiredFields.Total; i++)
            {
                int col = _colNumsOfRequiredFields[(int)(RequiredFields)i];
                RG rg = ws.Range[ws.Cells[firstRow, col], ws.Cells[lastRow, col]];
                _2DArrayOf1DData = (object[,])rg.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                
                for (int r = 0; r < qRows; r++)
                {
                    _objectArray[r, i] = _2DArrayOf1DData[(r+1),dummyColumnIndex];
                }
            }

            for (int r = 0; r < qRows; r++)
            {
                PODictionaryInExpRep po = new PODictionaryInExpRep()
                {
                    PONum = Convert.ToString(_objectArray[r, (int)RequiredFields.PoNumberCol]),
                    POLineNum = Convert.ToDouble(_objectArray[r, (int)RequiredFields.LineNumberCol]),
                };
                if (!_poDictionaryInExpRep.ContainsKey(Key))
                {
                    po.ExpRepXLLineNum = (r + firstRow);
                    po.MostRecentRevisedDeliveryDate = KAXL.ReadDateTime(_objectArray[r, (int)RequiredFields.RevDateCol]);
                    po.Status = new Status() { ExpRepStatus = Convert.ToString(_objectArray[r, (int)RequiredFields.StatusCol]) };
                    po.IsReceivedDatePresent = _objectArray[r, (int)RequiredFields.RecDateCol] != null ? true : false;
                }
            }
            //object[][,] _objectJaggedArray = new object[(int)RequiredFields.Total][,];

            //for (int i = 0; i < (int)RequiredFields.Total; i++)
            //{
            //    int col = _colNumsOfRequiredFields[(int)(RequiredFields)i];

            //    RG rg = ws.Range[ws.Cells[firstRow, col], ws.Cells[lastRow, col]];

            //    // this method only returns a 2D array, even it if is only 1D?
            //    object[,] _2DArrayOf1DColumn = (object[,])rg.get_Value(XlRangeValueDataType.xlRangeValueDefault);                
            //    _objectJaggedArray[i] = _2DArrayOf1DColumn;
            //}

            //for (int i = 0; i < qRows; i++)
            //{
            //    PODictionaryInExpRep po = new PODictionaryInExpRep()
            //    {
            //        PONum = Convert.ToString(_objectArray[(int)RequiredFields.PoNumberCol,dummyColumnIndex]),
            //        POLineNum = Convert.ToDouble(_objectArray[(int)RequiredFields.LineNumberCol,dummyColumnIndex]),
            //    };
            //    if (!_poDictionaryInExpRep.ContainsKey(Key))
            //    {
            //        po.ExpRepXLLineNum = (i + firstRow);
            //        po.MostRecentRevisedDeliveryDate = KAXL.ReadDateTime(_objectArray[(int)RequiredFields.RevDateCol,dummyColumnIndex]);
            //        po.Status = new Status() { ExpRepStatus = Convert.ToString(_objectArray[(int)RequiredFields.StatusCol,dummyColumnIndex]) };
            //        po.IsReceivedDatePresent = _objectArray[(int)RequiredFields.RecDateCol,dummyColumnIndex] != null ? true : false;
            //    }                                       
            //}
        }

        private List<int> LoadListOfColNumsOfReqFields()
        {
            List<int> listOfColNumsOfRequiredFields = new List<int>((int)RequiredFields.Total)
            {
                m.ExpRepColumn.PONumber,
                m.ExpRepColumn.LineNumber,
                m.ExpRepColumn.RecDate,
                m.ExpRepColumn.RevisedSchedDelDate,
                m.ExpRepColumn.Status
            };

            return listOfColNumsOfRequiredFields;
        }

        private void LoadDictionaryWithPOs()
        {
            m.kaxlApp.WS = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];

            WS ws = m.kaxlApp.WS;
            var k = m.kaxlApp.KAXL_RG;
            k = new KAXLApp.KAXLRange(m.kaxlApp, RangeType.WorkSheet);
                       
            m.kaxlApp.ErrorTracker.ProgramStage = "Reading Expedite Report";

            string poNum, key;
            double lineNum;
            DateTime receivedDate, revDate;
            bool isReceivedDatePresent;

            for (int r = 3; r < k.Row.End; r++)
            {
                m.kaxlApp.ErrorTracker.Row = r;

                try
                {
                    poNum = ws.Cells[r, m.ExpRepColumn.PONumber].Value2;
                    lineNum = Math.Round((double)ws.Cells[r, m.ExpRepColumn.LineNumber].Value2);
                    key = poNum + Convert.ToString(lineNum);                    

                    if (_poDictionaryInExpRep.ContainsKey(key))
                    {
                        key = null;
                    }
                    else
                    {
                        revDate = KAXL.ReadDateTime(ws.Cells[r, m.ExpRepColumn.RevisedSchedDelDate].Value);
                        receivedDate = KAXL.ReadDateTime(ws.Cells[r, m.ExpRepColumn.RecDate].Value);
                        isReceivedDatePresent = (ws.Cells[r, m.ExpRepColumn.RecDate].Value is null) ? false : true;

                        Status status = new Status()
                        {
                            ExpRepStatus = ws.Cells[r, m.ExpRepColumn.Status].Value2,
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

                        string itemNum = Convert.ToString(ws.Cells[r, m.ExpRepColumn.ItemNumber].Value2);
                        string vendName = ws.Cells[r, m.ExpRepColumn.VendorName].Value2;

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
                catch
                {
                    m.kaxlApp.ErrorTracker.AddNewError("Expedite Report Reading error - Row #" + Convert.ToString(r));
                }
            }

            //try
            //{
            //    for (int r = KAXL.FindFirstRowAfterHeader(ws); r < k.Row.End; r++)
            //    {
            //        m.kaxlApp.ErrorTracker.Row = r;

            //        poNum = (string)k[r, m.ExpRepColumn.PONumber];
            //        lineNum = Math.Round((double)k[r, m.ExpRepColumn.LineNumber]);
            //        key = poNum + Convert.ToString(lineNum);
            //        DateTime receivedDate = KAXL.ReadDateTime(k[r, m.ExpRepColumn.RecDate]);
            //        IsReceivedDatePresent = receivedDate == DateTime.MinValue ? false : true;

            //        if (_poDictionaryInExpRep.ContainsKey(key))
            //        {
            //            key = null;
            //        }
            //        else
            //        {
            //            DateTime revDate = KAXL.ReadDateTime(k[r, m.ExpRepColumn.RevisedSchedDelDate]);
            //            bool isReceivedDatePresent = (k[r, m.ExpRepColumn.RecDate] is null) ? false : true;

            //            Status status = new Status()
            //            {
            //                ExpRepStatus = (string)k[r, m.ExpRepColumn.Status],
            //                PONum = poNum,
            //            };

            //            _poDictionaryInExpRep.Add(key, new PODictionaryInExpRep()
            //            {
            //                ExpRepXLLineNum = r,
            //                IsReceivedDatePresent = isReceivedDatePresent,
            //                MostRecentRevisedDeliveryDate = revDate,
            //                POLineNum = lineNum,
            //                PONum = poNum,
            //                Status = status
            //            });

            //            string itemNum = Convert.ToString(k[r, m.ExpRepColumn.ItemNumber]);
            //            string vendName = (string)k[r, m.ExpRepColumn.VendorName];

            //            if (itemNum != null && (string)k[r, m.ExpRepColumn.ItemDescription] == null)
            //            {
            //                // do this if there is an item number, but not description
            //                Item i = m.ItemDict[itemNum];
            //                ws.Cells[r, m.ExpRepColumn.ItemDescription].Value2 = i.Desc;
            //            }

            //            // check if vendor name is present, and add to update list if not.
            //            if (vendName == null)
            //            {
            //                ws.Cells[r, m.ExpRepColumn.VendorName].Value2 = m.VendorDict[vendName];
            //            }
            //        }
            //    }
            //}
            //catch
            //{
            //    m.kaxlApp.ErrorTracker.AddNewError("Reading ExpRep, Line #" + Convert.ToString(m.kaxlApp.ErrorTracker.Row));
            //}
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
