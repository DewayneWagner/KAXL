using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EXPREP_V2;

namespace DKARibbon.BOM_Scrubber
{
    public class BOM
    {
        public BOM() { }
        public List<Item> ItemList { get; set; } = new List<Item>();
        public bool IsTopLevel { get; set; }
        public bool IsLowestLevel { get; set; }
        public BOMNumber BOMNumber { get; set; }
        public string Name { get; set; }
        public int QTYInParent { get; set; }
        public int QSubs { get; set; }
        public bool ItemIsBOM { get; set; }
        public bool IsNewestVersionOfBOM { get; set; }
    }
}
