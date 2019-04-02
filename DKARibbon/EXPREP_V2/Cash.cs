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

        public Cash() { }

        public Cash(Master m) => M = m;

        public Cash(string cur,double netAmount, DateTime createdDate, Master m, double quantity)
        {
            M = m;
            Currency = cur;
            int year = createdDate.Year;
            int month = createdDate.Month;
            NetAmount = netAmount;
            
            UnitPriceCAD = (M.ExRateDict[cur, "CAD", year, month] * netAmount) / quantity; 
            UnitPriceUSD = (M.ExRateDict[cur, "USD", year, month] * netAmount) / quantity;

            CAD = UnitPriceCAD * quantity;
            USD = UnitPriceUSD * quantity;
        }

        public double CAD { get; set; }
        public double USD { get; set; }
        public double UnitPriceCAD { get; set; }
        public double UnitPriceUSD { get; set; }

        public string Currency { get; set; }
        public double NetAmount { get; set; }
        public static Cash ZeroedOutCash() => new Cash() { NetAmount = 0, CAD = 0, USD = 0, UnitPriceCAD = 0, UnitPriceUSD = 0 };

    }    
}
