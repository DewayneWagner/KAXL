using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DKARibbon.BOM_Scrubber
{
    public class BOMDictionary
    {
        private Dictionary<string, BOM> _bomDictionary;
        public BOMDictionary()
        {
            _bomDictionary = new Dictionary<string, BOM>();
        }
        public BOM this[string key]
        {
            get => _bomDictionary[key];
            set => _bomDictionary[key] = value;
        }
    }
}
