using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DKARibbon.BOM_Scrubber
{
    public class BOMNumber
    {
        private string _short;
        private int _version;
        public BOMNumber(string rawBOMNumber) { Long = rawBOMNumber; }
        public string Long { get; set; }
        public string Short
        {
            get => _short;
            set => _short = Convert.ToString(Long.ToCharArray(1, Long.Length - 4));
        }
        public int Version
        {
            get => _version;
            set
            {
                char[] cA = Long.ToCharArray(Long.Length, 1);
                string v = cA.ToString();
                int version = 0;
                bool successful = Int32.TryParse(v, out version);
                _version = (successful) ? version : 0;
            }
        }
    }
}
