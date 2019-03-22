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

        public Cash() { }

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

        private double _cad;
        public double CAD
        {
            get => _cad;
            set
            {
                try { _cad = M.ExRateDict[Currency, "CAD", Year, Month] * NetAmount; }
                catch { _cad = 0; }
            }            
        }

        private double _usd;
        public double USD
        {
            get => _usd;
            set
            {
                try { _usd = M.ExRateDict[Currency, "USD", Year, Month] * NetAmount; }
                catch { _usd = 0; }
            }
        }

        private double _unitPriceCAD;
        public double UnitPriceCAD
        {
            get => _unitPriceCAD;
            set
            {
                try { _unitPriceCAD = M.ExRateDict[Currency, "CAD", Year, Month] * UnitPrice; }
                catch { _unitPriceCAD = 0; }
            }
        }

        public double _unitPriceUSD;
        public double UnitPriceUSD
        {
            get => _unitPriceUSD;
            set
            {
                try { _unitPriceUSD = M.ExRateDict[Currency, "USD", Year, Month] * UnitPrice; }
                catch { _unitPriceUSD = 0; }
            }
        }        
        public string Currency { get; set; }        
        public double NetAmount { get; set; }

        public static Cash ZeroedOutCash()
        {
            return new Cash()
            {
                NetAmount = 0,
                CAD = 0,
                UnitPriceCAD = 0,
                UnitPriceUSD = 0,
                USD = 0,
            };
        }
    }    
}
