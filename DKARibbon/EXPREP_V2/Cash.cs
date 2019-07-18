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
        public Cash() { }

        public Cash(string cur,double netAmount, DateTime createdDate, Master m, double quantity)
        {
            Currency = cur;
            int year = createdDate.Year;
            int month = createdDate.Month;
            NetAmount = netAmount;
            
            UnitPriceUSD = (m.ExRateDict[cur, "USD", year, month] * netAmount) / quantity;

            USD = UnitPriceUSD * quantity;
        }

        public double USD { get; set; }
        public double UnitPriceUSD { get; set; }
        public string Currency { get; set; }
        public double NetAmount { get; set; }
        public static Cash ZeroedOutCash() => new Cash() { NetAmount = 0, USD = 0, UnitPriceUSD = 0 };

    }    
}
