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
    //public static class VendorUpdate
    //{
    //    public static void VendorM(WS ws, RG rg)
    //    {
    //        int endRow = rg.Row + rg.Rows.Count - 1;
    //        int endCol = rg.Column + rg.Columns.Count - 1;

    //        string VendorNum = null;
    //        string VendorName = null;

    //        Dictionary<string, Vendor> VendorDict = new Dictionary<string, Vendor>();

    //        for (int r = rg.Row; r <= endRow; r++)
    //        {
    //            VendorName = ws.Cells[r, rg.Column].Value2;
    //            VendorNum = ws.Cells[r, (rg.Column + 1)].Value2;

    //            if (!VendorDict.ContainsKey(VendorName))
    //            {
    //                Vendor vendor = new Vendor(VendorName, VendorNum);
    //                VendorDict.Add(VendorName, vendor);
    //            }
    //        }

    //        XL xl = new XL();
    //        Workbook wb = xl.Workbooks.Open(@"R:\Supply Chain\ExpediteReport\ExpediteReport.xlsx");
    //        KAXL.CloseAndSaveWBIfOpen(wb);
    //        Worksheet expRep = wb.Sheets[1];

    //        if (ws.AutoFilter != null)
    //            ws.AutoFilterMode = false;

    //        DestColID dCol = new DestColID(expRep);

    //        int LR = KAXL.LastRow(expRep, 1);
    //        VendorNum = null;

    //        for (int r = 1; r <= LR; r++)
    //        {
    //            var val = expRep.Cells[r, dCol.ItemNumber].Value2;

    //            if (val is double)
    //                VendorNum = Convert.ToString(val);
    //            else
    //                VendorNum = val;

    //            if (VendorNum != null && VendorDict.ContainsKey(VendorNum))
    //            {
    //                Vendor vendor = VendorDict[VendorNum];
    //                expRep.Cells[r, dCol.VendorName].Value2 = vendor.VendorCName;
    //                expRep.Cells[r, dCol.VendorAccount].Value2 = vendor.VendorCAccount;
    //            }
    //        }
    //        MessageBox.Show("Finished updating Vendor#'s.");
    //        KAXL.CloseApp(xl);
    //    }
    //}
}
