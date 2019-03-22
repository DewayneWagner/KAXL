using System.Collections.Generic;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using DKAExcelStuff;
using MDC = EXPREP_V2.Master.MasterDataColumnsE;

namespace EXPREP_V2
{
    public class Vendor
    {
        Master m;
        private readonly Dictionary<string, Vendor> _vendorDictionary;
        private readonly List<string> _vendorNamesNotInDictionary;
        private enum VendorColumnOrder { Nada, Name, Number }

        public Vendor() { }

        public Vendor(Master master)
        {
            m = master;
            _vendorDictionary = new Dictionary<string, Vendor>();
            _vendorNamesNotInDictionary = new List<string>();
            LoadVendorDictionary();
        }

        public string Code { get; set; }
        public string Name { get; set; }

        private void LoadVendorDictionary()
        {
            m.kaxlApp.WS = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.MasterData];

            WS ws = m.kaxlApp.WS;
            var k = m.kaxlApp.KAXL_RG;

            m.kaxlApp.RG = ws.Range[ws.Cells[KAXL.FindFirstRowAfterHeader(ws), (int)MDC.VendorName],
                ws.Cells[KAXL.LastRow(ws, (int)MDC.VendorName), (int)MDC.VendorAccount]];
            k = new KAXLApp.KAXLRange(m.kaxlApp, RangeType.CodedRangeSetKAXLAppRG);
            string name;

            for (int r = 1; r < k.Row.End; r++)
            {
                name = (string)k[r, (int)VendorColumnOrder.Name];

                if (_vendorDictionary.ContainsKey(name))
                {
                    name = null;
                }
                else
                {
                    _vendorDictionary.Add(name, new Vendor()
                    {
                        Name = name,
                        Code = (string)k[r, (int)VendorColumnOrder.Number],
                    });
                }                
            }
        }

        public Vendor this[string key] => key != null && _vendorDictionary.ContainsKey(key) ? _vendorDictionary[key] : AddToVendorNumbersThatArentInDict(key);

        private Vendor AddToVendorNumbersThatArentInDict(string key)
        {
            _vendorNamesNotInDictionary.Add(key);
            return new Vendor()
            {
                Name = key,
                Code = null,
            };
        }
        public List<string> VendorNumbersThatArentInDictL() => _vendorNamesNotInDictionary;
        public bool IsVendorNumbersThatArentInDict() => (_vendorNamesNotInDictionary.Count > 0) ? true : false;
    }    
}
