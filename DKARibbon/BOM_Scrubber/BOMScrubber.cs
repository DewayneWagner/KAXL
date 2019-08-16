using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DKAExcelStuff;
using EXPREP_V2;
using static DKAExcelStuff.KAXLApp;

namespace DKARibbon.BOM_Scrubber
{
    public class BOMScrubber
    {
        public static int maxBOMLevels = 10;
        private enum Index { ItemNum, ItemName, BOMNum, BOMName, QTYPerSeries, QTYItemPerBom, Total }

        private KAXLApp kaxlApp;
        public KAXLRange k;
        private Dictionary<string, BOM> _individualBOMDictionary; // list of all BOMs
        private Stack<BOM> _allBOMsWithHierarchyStack; // 
        private List<string> _allBOMNumbers;
        public BOMScrubber(KAXLApp ka)
        {
            kaxlApp = ka;
            _individualBOMDictionary = new Dictionary<string, BOM>();
            _allBOMsWithHierarchyStack = new Stack<BOM>();

            LoadRawData();
            ColID = new BOMColID(LoadColumnHeadingList());
            _allBOMNumbers = LoadListOfAllBOMs();

            BuildListOfBOMs();


            ExpRepData = new ExpRepData();
        }

        private BOMColID ColID { get; set; }

        private void BuildListOfBOMSWithHierarchy()
        {
            int i = 0;

            foreach (KeyValuePair<string,BOM> b in _individualBOMDictionary)
            {
                
            }
        }

        private void BuildListOfBOMs()
        {
            string key = null;
            for (int i = 1; i <= k.Row.End; i++)
            {
                key = Convert.ToString(k[i, ]);
                _individualBOMDictionary.Add(key,new BOM()
                {
                    IsTopLevel = DetermineIfIsTopLevel(i),
                    Number = Convert.ToString(k[i, ColID.BOMNumber]),
                    //Name = Convert.ToString(k[i, Col.BOMName]),
                    QTYInParent = Convert.ToInt32(k[i, ColID.QuantityOfItemInParent]),
                    QSubs = GetQSubs(i)
                });
            }
        }

        private bool DetermineIfIsTopLevel(int bomRow)
        {
            var bom = k[bomRow, ColID.BOMNumber];
            for (int i = 1; i <= k.Row.End; i++)
            {
                if (k[i, ColID.ItemNumber] == bom)
                    return false;
            }
            return true;
        }
        //private int GetQSubs(int bomRow)
        //{
        //    var bom = k[bomRow, ColID.BOMNumber];
        //    bool isAssembly;
        //    int q = 0;
        //    for (int i = 1; i <= k.Row.End; i++)
        //    {
        //        isAssembly = Convert.ToBoolean(k[i, Col.IsAssembly]);

        //        if (k[i, ColID.BOMNumber] == bom && isAssembly)
        //            q++;
        //    }
        //    return q;
        //}

        public ExpRepData ExpRepData { get; set; }

        private void LoadRawData()
        {
            int LR = KAXL.LastRow(kaxlApp.WS, 1);
            int LC = 18;

            kaxlApp.RG = kaxlApp.WS.Range[kaxlApp.WS.Cells[1, 1], kaxlApp.WS.Cells[LR, LC]];
            k = new KAXLApp.KAXLRange(kaxlApp, RangeType.CodedRangeSetKAXLAppRG);
        }

        private List<string> LoadColumnHeadingList()
        {
            List<string> _headings = new List<string>();

            for (int i = 1; i <= k.Col.Q; i++)
            {
                _headings.Add(Convert.ToString(k[1, i]));
            }

            return _headings;
        }

        private List<string> LoadListOfAllBOMs()
        {
            List<string> _allBOMs = new List<string>();

            for (int i = 1; i <= k.Row.Q; i++)
            {
                _allBOMs.Add(TrimEndOffBOM(Convert.ToString(k[i, ColID.BOMNumber])));
            }
            return _allBOMs;
        }
        
        private bool DetermineIfItemIsBOM(string bomNumber) => _allBOMNumbers.Contains(bomNumber) ? true : false;
        private class TopLevelBOM : BOM
        {
            private List<BOM> _BOMHierarchy;
            public TopLevelBOM()
            {
                _BOMHierarchy = new List<BOM>(maxBOMLevels);
            }
            public BOM this[int level]
            {
                get => _BOMHierarchy[level];
                set => _BOMHierarchy[level] = value;
            }
        }
    }
}
