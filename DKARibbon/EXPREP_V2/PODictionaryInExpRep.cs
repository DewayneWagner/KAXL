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

            string poNum, key;
            double lineNum;

            for (int r = 0; r < qRows; r++)
            {
                poNum = Convert.ToString(_objectArray[r, (int)RequiredFields.PoNumberCol]);
                lineNum = Convert.ToDouble(_objectArray[r, (int)RequiredFields.LineNumberCol]);
                key = GetKey(poNum, lineNum);

                PODictionaryInExpRep po = new PODictionaryInExpRep()
                {
                    PONum = poNum,
                    POLineNum = lineNum,
                };
                if (!IsDuplicate(key))
                {
                    po.ExpRepXLLineNum = (r + firstRow);                    
                    po.MostRecentRevisedDeliveryDate = KAXL.ReadDateTime(_objectArray[r, (int)RequiredFields.RevDateCol]);
                    po.Status = new Status() { ExpRepStatus = Convert.ToString(_objectArray[r, (int)RequiredFields.StatusCol]) };
                    po.IsReceivedDatePresent = _objectArray[r, (int)RequiredFields.RecDateCol] != null ? true : false;
                    _poDictionaryInExpRep[key] = po;                    
                }                
            }
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

        public bool IsDuplicate(string poNum, double lineNum) =>
            (_poDictionaryInExpRep.ContainsKey(poNum + Convert.ToString(Math.Floor(lineNum)))) ? true : false;

        public bool IsDuplicate(string key) => key != null && _poDictionaryInExpRep.ContainsKey(key) ? true : false;

        public bool ContainsKey(string key) => _poDictionaryInExpRep.ContainsKey(key);

        public string GetKey(string poNum, double lineNum) => poNum + Convert.ToString(Math.Floor(lineNum));

        public PODictionaryInExpRep this[string key]
        {
            get => key != null && _poDictionaryInExpRep.ContainsKey(key) ? _poDictionaryInExpRep[key] : null;
            set => _poDictionaryInExpRep[key] = value;
        }
        
    }
}
