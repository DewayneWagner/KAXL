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
        private ColIDL colID;

        public BOMColID(KAXLApp k)
        {
            colID = new ColIDL(k.WS);
        }
        public int ItemNum => colID.GetColNum("Item number");
        public int ItemName => colID.GetColNum("Product name");

    }
}
