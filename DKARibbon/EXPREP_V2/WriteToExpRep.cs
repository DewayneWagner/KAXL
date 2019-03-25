using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using DKAExcelStuff;
using System.Runtime.InteropServices;

namespace EXPREP_V2
{
    public class WriteToExpRep
    {
        private object[,] _scrubbedPOLinesObjArray;
        private List<ScrubbedPOLine> _scrubbedPOList;
        Master m;
        private int rowQ, colQ;

        public WriteToExpRep(Master master, List<ScrubbedPOLine> scrubbedPOList)
        {
            m = master;
            rowQ = scrubbedPOList.Count + 1;
            colQ = m.ExpRepColumn.totalColumnsInExpRep + 1;
            // +1 to correct for zero-based array - all 0 indexes don't exist

            _scrubbedPOLinesObjArray = new object[rowQ, colQ];
            _scrubbedPOList = scrubbedPOList;

            LoadObjArray();
            WriteArrayInExpRep();
        }

        public object this[int r, int c]
        {
            get => _scrubbedPOLinesObjArray[r,c];
            set => _scrubbedPOLinesObjArray[r, c] = value;
        }

        private void LoadObjArray()
        {
            for (int i = 1; i < rowQ; i++)
            {
                i--; // to account for 0 index in list, but not in array
                ScrubbedPOLine po = _scrubbedPOList[i];

                i++;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.AttentionInfo] = po.Source.OriginalAttentionInfo;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.CAD] = po.Cash.CAD;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Category] = po.Category.CleanCategory;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Createdby] = po.Source.CreatedBy;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Curr] = po.Cash.Currency;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.DateAdded] = DateTime.Today;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Direct] = po.Direct;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Entity] = po.Entity;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ExpediteRequired] = po.ExpediteRequired;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Expeditor] = po.Source.CreatedBy;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ICO] = po.ICO;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ItemDescription] = po.ItemX.Desc;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.ItemNumber] = po.ItemX.Num;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.LineNumber] = po.LineNumber;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Month] = po.Dates.Month;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.NetAmount] = po.Cash.NetAmount;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.OriginalSchedDelDate] = po.Dates.OriginalScheduledDelivery;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.POCreatedDate] = po.Dates.POCreated;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.PONumber] = po.PONum;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.POSourceCode] = po.Source.Code;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.POSourceType] = po.Source.Type;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Quantity] = po.Quantity;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Quarter] = po.Dates.Quarter;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.RecDate] = po.Dates.Received;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Receiver] = po.Receiver;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Requester] = po.Source.Requester;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.RevisedSchedDelDate] = po.Dates.RevisedScheduledDeliveryDate;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Status] = po.Status.CleanStatus;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.UnitPriceCAD] = po.Cash.UnitPriceCAD;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.UnitPriceUSD] = po.Cash.UnitPriceUSD;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.USD] = po.Cash.USD;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.VendorAccount] = po.Vendor.Code;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.VendorName] = po.Vendor.Name;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.WH] = po.WH;
                _scrubbedPOLinesObjArray[i, m.ExpRepColumn.Year] = po.Dates.Year;
            }
        }
        public void WriteArrayInExpRep()
        {
            //m.kaxlApp.WS = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            //WS ws = m.kaxlApp.WS;

            WS ws = m.kaxlApp.WB.Sheets[(int)Master.SheetNamesE.ExpRep];
            RG rg = ws.Range[ws.Cells[KAXL.FindFirstRowAfterHeader(ws), 1], ws.Cells[KAXL.LastRow(ws, 1) + 1, colQ]];
            //rg.set_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault, _scrubbedPOLinesObjArray);

            //Microsoft.Office.Tools.Excel.NamedRange nr = ws.Range[ws.Cells[KAXL.FindFirstRowAfterHeader(ws), 1], ws.Cells[KAXL.LastRow(ws, 1) + 1, colQ]];
            //nr.set_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault, _scrubbedPOLinesObjArray);

            Microsoft.Office.Tools.Excel.NamedRange nr = ws.Range[ws.Cells[KAXL.FindFirstRowAfterHeader(ws), 1], ws.Cells[KAXL.LastRow(ws, 1) + 1, colQ]];            

            Marshal.ReleaseComObject(nr);
        }
    }
}
