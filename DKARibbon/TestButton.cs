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
            KAXLTest k = new KAXLTest();
            List<string> headingsList = k.GetHeadingsList(kaxlApp.WS);

            
        }
        
    }
    public class KAXLTest
    {
        public List<string> GetHeadingsList(WS ws)
        {
            int headersRow = GetHeadingsRow();
            int lastCol = LastCol(ws, headersRow);

            RG headingsRange = ws.Range[ws.Cells[headersRow, 1], ws.Cells[headersRow, lastCol]];

            object[,] _headingsArray = Get2DObjectArray(headingsRange);

            return GetHeadingsList();

            int GetHeadingsRow()
            {
                int FR = 0;
                do { FR++; } while (ws.Cells[FR, 1].Value2 == null);
                return FR;
            }
            List<string> GetHeadingsList()
            {
                List<string> headingsList = new List<string>();

                for (int i = 1; i < _headingsArray.GetLength(1); i++)
                {
                    headingsList.Add(Convert.ToString(_headingsArray[1,i]));
                }

                return headingsList;
            }
        }
        public object[] Get1DObjectArray(RG rg) => (object[])rg.get_Value(XlRangeValueDataType.xlRangeValueDefault);
        public object[,] Get2DObjectArray(RG rg) => (object[,])rg.get_Value(XlRangeValueDataType.xlRangeValueDefault);
        public int LastCol(WS ws, int headerRow)
        {
            bool foundEmptyCol = false;
            int searchCol = 1;

            do
            {
                if (ws.Cells[headerRow, searchCol].Value2 == null) { return (searchCol - 1); }
                searchCol++;
            } while (!foundEmptyCol);
            return 1;
        }
    }
}