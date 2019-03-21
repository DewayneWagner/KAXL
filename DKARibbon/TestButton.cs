using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Microsoft.Office.Interop.Excel;
using XL = Microsoft.Office.Interop.Excel.Application;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;

using DKARibbon;
using System.Net;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Net.Http;

namespace DKAExcelStuff
{
    public class TestButton
    {
        public static void TestM(KAXLApp kaxlApp)
        {
            int rowQ = kaxlApp.WorkSheetRange.Row.Q;
            int startRow = kaxlApp.WorkSheetRange.Row.Start;
            int endRow = kaxlApp.WorkSheetRange.Row.End;

            int startCol = kaxlApp.WorkSheetRange.Col.Start;
            int endCol = kaxlApp.WorkSheetRange.Col.End;
            int colQ = kaxlApp.WorkSheetRange.Col.Q;
            
            object[,] testArray = new object[kaxlApp.WorkSheetRange.Row.Q + 1, kaxlApp.WorkSheetRange.Col.Q + 1];

            for (int r = 1; r <= kaxlApp.WorkSheetRange.Row.Q; r++)
            {
                for (int c = 1; c <= kaxlApp.WorkSheetRange.Col.Q; c++)
                {
                    testArray[r, c] = kaxlApp.WorkSheetRange[r, c];
                }
            }
        }
    }  
}