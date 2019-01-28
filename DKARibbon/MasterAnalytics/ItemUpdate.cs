using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using XL = Microsoft.Office.Interop.Excel.Application;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using System.Windows.Forms;

namespace DKAExcelStuff
{
    public static class ItemUpdate
    {
        //public static void ItemUpdateM(WS ws, RG rg)
        //{
        //    int endRow = rg.Row + rg.Rows.Count - 1;
        //    int endCol = rg.Column + rg.Columns.Count - 1;

        //    string itemN = null;
        //    string itemD = null;
        //    string itemC = null;

        //    Dictionary<string, ItemDesc> itemDict = new Dictionary<string, ItemDesc>();

        //    for (int r = rg.Row; r <= endRow; r++)
        //    {
        //        var val = ws.Cells[r, rg.Column].Value2;

        //        if (val is double)
        //            itemN = Convert.ToString(val);
        //        else
        //            itemN = val;

        //        if (!itemDict.ContainsKey(itemN))
        //        {
        //            itemD = ws.Cells[r, (rg.Column + 1)].Value2;
        //            itemC = ws.Cells[r, endCol].Value2;

        //            ItemDesc item = new ItemDesc(itemN, itemD, itemC);

        //            itemDict.Add(itemN, item);
        //        }
        //    }

        //    XL xl = new XL();
        //    Workbook wb = xl.Workbooks.Open(@"R:\Supply Chain\ExpediteReport\ExpediteReport.xlsx");
        //    KAXL.CloseAndSaveWBIfOpen(wb);
        //    Worksheet expRep = wb.Sheets[1];

        //    if (ws.AutoFilter != null)
        //        ws.AutoFilterMode = false;

        //    DestColID dCol = new DestColID(expRep);

        //    int LR = KAXL.LastRow(expRep, 1);
        //    itemN = null;

        //    for (int r = 1; r <= LR; r++)
        //    {
        //        var val = expRep.Cells[r, dCol.ItemNumber].Value2;

        //        if (val is double)
        //            itemN = Convert.ToString(val);
        //        else
        //            itemN = val;

        //        if (itemN != null && itemDict.ContainsKey(itemN))
        //        {
        //            ItemDesc item = itemDict[itemN];
        //            expRep.Cells[r, dCol.ItemDescription].Value2 = item.ItemD;
        //            expRep.Cells[r, dCol.ItemCat].Value2 = item.ItemCategory;
        //        }
        //    }
        //    MessageBox.Show("Finished updating Item#'s.");
        //    KAXL.CloseApp(xl);
        //}
    }
}
