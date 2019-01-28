using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DKAExcelStuff;
using WS = Microsoft.Office.Interop.Excel.Worksheet;

namespace EXPREP_V2
{
    public class ExRate 
    {
        public ExRate() {}

        public ExRate(string key,double exrate)
        {
            Key = key;
            exRate = exrate;
        }

        public ExRate(string currFrom,string currTo,int year,int month,double exrate)
        {
            CurrFrom = currFrom;
            CurrTo = currTo;
            Year = year;
            Month = month;
            exRate = exrate;
        }

        public string Key { get; set; }

        public string GetKey(string currFrom, string currTo, int year, int month) => (currFrom + currTo + year + month);

        private string CurrFrom { get; set; }
        private string CurrTo { get; set; }
        private int Year { get; set; }
        private int Month { get; set; }
        public double exRate { get; set; }
    }
    public class ExRateDict  
    {
        enum ExRateColE {Nada,Key,CurrFrom,CurrTo,Year,Month,ExRate}
        private readonly Dictionary<string, double> exRateDict;

        public ExRateDict() {}

        public ExRateDict(KAXLApp kaxlApp)
        {
            exRateDict = new Dictionary<string, double>();
            APP = kaxlApp;
            LoadDict();
        }
        private KAXLApp APP { get; set; }

        private void LoadDict()
        {
            WS ws = APP.WB.Sheets[(int)Master.SheetNamesE.MasterData];
            int startRow = 2;
            int LR = KAXL.LastRow(ws, (int)ExRateColE.Key);
            string key;
            double exrate;

            for (int i = startRow; i < LR; i++)
            {
                key = ws.Cells[i, (int)ExRateColE.Key].Value2;
                try
                {
                    exrate = ws.Cells[i, (int)ExRateColE.ExRate].Value2;
                }
                catch
                {
                    exrate = 1;
                }
                
                ExRate x = new ExRate(key, exrate);
                exRateDict.Add(x.Key, x.exRate);
            }
        }
        public double this[string key]
        {
            get => key != null && exRateDict.ContainsKey(key) ? exRateDict[key] : 0;
            set => exRateDict[key] = value;
        }
        public double this[string currFrom,string currTo,int year, int month] 
        {
            get
            {
                string key = currFrom + currTo + year + month;
                return key != null && exRateDict.ContainsKey(key)? exRateDict[key] : 0;
            }            
        }
    }
}
