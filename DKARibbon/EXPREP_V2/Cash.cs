using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DKAExcelStuff;

namespace EXPREP_V2
{
    public class Cash
    {
        private Master M;
        private double UnitPrice;

        public Cash(Master m) => M = m;

        public Cash(string cur,string unitPrice,string netAmount, string createdDate, Master m)
        {
            M = m;
            Currency = cur;
            Year = KAXL.YearFromString(createdDate);
            Month = KAXL.MonthFromString(createdDate);
            UnitPrice = Convert.ToDouble(unitPrice);
            NetAmount = Convert.ToDouble(netAmount);
        }
        public Cash(bool zerodOutForMultiLinePOs)
        {
            NetAmount = 0;
            NetAmount = 0;
        }

        private int Year { get; set; }
        private int Month { get; set; }
                
        public double CAD
        {
            get
            {
                try { return M.ExRateDict[Currency, "CAD", Year, Month] * NetAmount; }
                catch { return 0; }
            }
        }
        public double USD
        {
            get
            {
                try { return M.ExRateDict[Currency, "USD", Year, Month] * NetAmount; }
                catch { return 0; }
            }
        }
        public double UnitPriceCAD
        {
            get
            {                
                try { return M.ExRateDict[Currency, "CAD", Year, Month] * UnitPrice;}
                catch { return 0; }
            }
        }
        public double UnitPriceUSD
        {
            get
            {
                try { return M.ExRateDict[Currency, "USD", Year, Month] * UnitPrice; }
                catch { return 0; }
            }
        }        
        public string Currency { get; set; }        
        private double NetAmount { get; set; }
    }    
}
