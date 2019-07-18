using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DKARibbon.BOM_Scrubber
{
    public class CleanBOMData
    {
        public string ItemNum { get; set; }
        public string ItemName { get; set; }
        public string BOMNum { get; set; }
        public string BOMName { get; set; }
        public int QtyItemInBOM { get; set; }
        public int QTY_BOM { get; set; }
        public string UOM { get; set; }
        public string MasterBOMNumber { get; set; }
        public string MasterBOMName { get; set; }
    }
}
