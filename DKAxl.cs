using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;

namespace DKAClassLibrary
{
    // This library will contain various code and methods that will be used frequently when using VSTO

    public class DKAxl
    {
        public DKAxl()
        {

        }
        public int LastRow(Worksheet ws)
        {
            int LR = 0;
            LR = ws.Rows.Count;
            return LR;
        }
        public int LastCol(Worksheet ws)
        {
            int LC = 0;
            LC = ws.Columns.Count;
            return LC;
        }
        public int NextRow(Worksheet ws)
        {
            int NR = 0;
            NR = ws.Rows.Count + 1;
            return NR;
        }
        public string[,] DataArr(Worksheet ws, int lr, int maxCleanCol, List<string> data)
        {
            string[,] dataArr = new string[lr, maxCleanCol];

            for (int r = 0; r < lr; r++)
            {
                for (int c = 0; c < maxCleanCol; c++)
                {
                    dataArr[r, c] = data[(r * c) + c];
                }
            }
            return dataArr;
        }


    }
}
