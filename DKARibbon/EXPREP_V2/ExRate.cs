using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DKAExcelStuff;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using MDC = EXPREP_V2.Master.MasterDataColumnsE;

namespace EXPREP_V2
{
    public class ExRate 
    {
        private readonly Dictionary<string, double> _exRateDictionary;
        Master m;

        public ExRate(Master master)
        {
            m = master;
            _exRateDictionary = new Dictionary<string, double>();
            LoadExRateDictionary();
        }

        private void LoadExRateDictionary()
        {
            m.kaxlApp.WS = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.MasterData];

            WS ws = m.kaxlApp.WS;
            var k = m.kaxlApp.KAXL_RG;

            int FirstRowOfData = KAXL.FindFirstRowAfterHeader(ws);
            int LR = KAXL.LastRow(ws, (int)MDC.ExRateKey);
            int LC = (int)MDC.ExRate;

            m.kaxlApp.RG = ws.Range[ws.Cells[FirstRowOfData, (int)MDC.ExRateKey],ws.Cells[LR, LC]];

            k = new KAXLApp.KAXLRange(m.kaxlApp, RangeType.CodedRangeSetKAXLAppRG);

            for (int r = 1; r < k.Row.End; r++)
            {
                _exRateDictionary.Add((string)k[r, (int)MDC.ExRateKey], (double)k[r, (int)MDC.ExRate]);
            }
        }

        public ExRate() {}

        public double this[string key]
        {
            get => _exRateDictionary[key];
            set => _exRateDictionary[key] = value;
        }
        public double this[string currFrom, string currTo, int year, int month]
        {
            get => _exRateDictionary[GetKey(currFrom, currTo, year, month)];
            set => _exRateDictionary[GetKey(currFrom, currTo, year, month)] = value;
        }

        public string GetKey(string currFrom, string currTo, int year, int month) => (currFrom + currTo + year + month);
        public double exRate { get; set; }
    }    
}
