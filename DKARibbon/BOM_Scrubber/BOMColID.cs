using DKAExcelStuff;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DKARibbon.BOM_Scrubber
{
    public class BOMColID
    {
        private ColIDL ColID;

        public BOMColID(List<string> _headings)
        {
            ColID = new ColIDL(_headings);
        }
        public int ItemNumber => ColID.GetColNum("Item number");
        public int BOMNumber => ColID.GetColNum("BOM");
        public int QuantityOfItemInParent => ColID.GetColNum("Quantity");
        public int UOM => ColID.GetColNum("Unit");
        public int ItemDescription => ColID.GetColNum("Product name");
    }
}
