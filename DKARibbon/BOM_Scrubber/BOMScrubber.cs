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
        private enum LineType { ItemHeader, BOMHeader, Data, Noise }
        private enum Index { ItemNum, ItemName, BOMNum, BOMName, QTYPerSeries, QTYItemPerBom, Total }

        private KAXLApp kaxlApp;
        public KAXLRange k;
        private List<CleanBOMData> _cleanData;

        public BOMScrubber(KAXLApp ka)
        {
            kaxlApp = ka;
            _cleanData = new List<CleanBOMData>();


            LoadRawData();
            ProcessRawData();

            ExpRepData = new ExpRepData();
        }

        public ExpRepData ExpRepData { get; set; }
        private ListOfBOMArrays BOM_Arr { get; set; }

        private void LoadRawData()
        {
            int LR = KAXL.LastRow(kaxlApp.WS, 1);
            int LC = 18;

            kaxlApp.RG = kaxlApp.WS.Range[kaxlApp.WS.Cells[1, 1], kaxlApp.WS.Cells[LR, LC]];
            kaxlApp.RG.UnMerge();
            k = new KAXLApp.KAXLRange(kaxlApp, RangeType.CodedRangeSetKAXLAppRG);
        }

        private void ProcessRawData()
        {
            string masterBomNum = null, masterBOMName = null, bomNum = null, bomName = null;
            int qtyBOM = 0, currentRow = 1, qty = 0;
            bool isFirstItemHeader = true;

            CleanBOMData currentCleanBOMDataLine = new CleanBOMData();

            LineType lt;
            for (int r = 1; r < k.Row.End; r++)
            {
                // row 1 will always be highest level BOM

                lt = GetRowType(Convert.ToString(k[r,SourceCol.NumberItemAndBOM]));

                switch (lt)
                {
                    case LineType.BOMHeader:
                        ProcessBOMHeaderLine(r);
                        break;
                    case LineType.ItemHeader:
                        currentRow = ProcessItemHeaderLine(r);
                        r = currentRow - 1; // ProcessItemHeaderLine processes several lines - so counter needs to be reset.  -1 to account for autoincrement in for loop
                        break;
                    default:
                        break;
                }
            }
            int ProcessItemHeaderLine(int row) // return row to process for next  iteration to reset loop counter
            {
                row++;  // because header isn't relevant - only data immediately below the header
                if (isFirstItemHeader)
                {
                    masterBomNum = Convert.ToString(k[row,SourceCol.NumberItemAndBOM]);
                    masterBOMName = Convert.ToString(k[row, SourceCol.NameItemAndBOM]);
                    isFirstItemHeader = false;
                }
                else
                {
                    do
                    {
                        _cleanData.Add(new CleanBOMData()
                        {
                            BOMName = bomName,
                            BOMNum = bomNum,
                            ItemName = Convert.ToString(k[row, SourceCol.NameItemAndBOM]),
                            ItemNum = Convert.ToString(k[row, SourceCol.NumberItemAndBOM]),
                            QtyItemInBOM = (Int32.TryParse(Convert.ToString(k[row, SourceCol.QTY_ItemInBOM]), out qty)) ? qty : 0,
                            UOM = Convert.ToString(k[row, SourceCol.UOM]),
                            QTY_BOM = qtyBOM,
                            MasterBOMNumber = masterBomNum,
                            MasterBOMName = masterBOMName
                        });

                        if(row == k.Row.End) { break; }
                        else { row++; }

                        lt = GetRowType(Convert.ToString(k[row, SourceCol.NumberItemAndBOM]));
                    } while (lt == LineType.Data);
                }
                
                return row;
            }
            void ProcessBOMHeaderLine(int row) 
            {
                row++;

                bomNum = Convert.ToString(k[row, SourceCol.NumberItemAndBOM]);
                bomName = Convert.ToString(k[row, SourceCol.NameItemAndBOM]);
                
                int num;
                bool success = Int32.TryParse(Convert.ToString(k[row, SourceCol.QTY_BOM]), out num);
                qtyBOM = success ? num : 1;
            }
        }

        private LineType GetRowType(string firstValue)
        {
            switch (firstValue)
            {
                case "Item number":
                    return LineType.ItemHeader;
                case "BOM":
                    return LineType.BOMHeader;
                case "Lines":
                case "Pure Technologies Ltd.":
                case null:
                    return LineType.Noise;
                default:
                    return LineType.Data;
            }
        }
        private class SourceCol
        {
            public static int NumberItemAndBOM = 1;
            public static int NameItemAndBOM = 2;
            public static int QTY_BOM = 3;
            public static int QTY_ItemInBOM = 11;
            public static int QTY_PerSeries = 13;
            public static int UOM = 15;
        }
        
        private class BOM
        {
            public BOM() { }
            public string Number { get; set; }
            public string Name { get; set; }
            public int QTYInParent { get; set; }
            public int Level { get; set; }
            public bool IsTopLevel { get; set; }
            public bool HasSubLevels { get; set; }
            public List<SUBBOM> ListOfSubBOMs { get; set; }
        }
        private class SUBBOM : BOM { }
        private class ListOfBOMArrays
        {
            int maxBOMLevels = 6;
            BOM[] _bomA;
            public ListOfBOMArrays()
            {
                _bomA = new BOM[maxBOMLevels];
            }
        }
    }
}
