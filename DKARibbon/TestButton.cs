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
            KAXL.ScrubItemNumbers(kaxlApp);
        }
    }  
}