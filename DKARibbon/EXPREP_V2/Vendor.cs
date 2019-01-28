using System.Collections.Generic;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using DKAExcelStuff;

namespace EXPREP_V2
{
    public class Vendor
    {
        public Vendor() { }

        public Vendor(string code, string name)
        {
            Code = code;
            Name = name;
        }
        public string Code { get; set; }
        public string Name { get; set; }

    }
    public class VendorDict
    {
        private readonly Dictionary<string, Vendor> vendorDict;
        private List<string> vendorNumbersThatArentInDictL;

        public VendorDict() { }

        private Master M;
        private int LR, startRow = 2;

        public VendorDict(Master m)
        {
            M = m;
            vendorDict = new Dictionary<string, Vendor>();
            LoadDict();
            vendorNumbersThatArentInDictL = new List<string>();
        }

        private void LoadDict()
        {
            WS ws = M.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.MasterData];
            LR = KAXL.LastRow(ws, (int)Master.MasterDataColumnsE.VendorAccount);

            for (int i = startRow; i <= LR; i++)
            {
                Vendor v = new Vendor(ws.Cells[i, (int)Master.MasterDataColumnsE.VendorAccount].Value2, 
                    ws.Cells[i, (int)Master.MasterDataColumnsE.VendorName].Value2);
                if (!vendorDict.ContainsKey(v.Code))
                    vendorDict.Add(v.Code, v);
            }
        }
        public Vendor this[string key]
        {
            get => key != null && vendorDict.ContainsKey(key) ? vendorDict[key] : AddTooVendorNumbersThatArentInDictL(key);
            set => vendorDict[key] = value;
        }

        private Vendor AddTooVendorNumbersThatArentInDictL(string key)
        {
            vendorNumbersThatArentInDictL.Add(key);

            string code = (key != null) ? key : "Code Missing";
            string name = "Name not on list";
            
            return new Vendor(code, name);
        }

        public List<string> VendorNumbersThatArentInDictL() => vendorNumbersThatArentInDictL;
        public bool IsVendorNumbersThatArentInDict() => (vendorNumbersThatArentInDictL.Count > 0) ? true : false;
    }
}
