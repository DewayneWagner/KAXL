﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Tools.Ribbon;
using DKARibbon;
using System.ComponentModel;
using SD = System.Data;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

namespace DKAExcelStuff
{
    // This library will contain various code and methods that will be used frequently when using VSTO

    public class KAXL
    {   
        //public static int LastRow(Worksheet ws, int col)
        //{
        //    string val = null;
        //    bool done = false;

        //    List<int> incrementsL = new List<int>() { 10000, 1000, 100, 1 };
        //    int length = incrementsL.Count;

        //    int LR = incrementsL[0] + 1;

        //    for (int i = 0; i < length; i++)
        //    {
        //        do
        //        {
        //            var value = ws.Cells[LR, col].Value2;

        //            if (value is double)
        //                val = Convert.ToString(value);
        //            else
        //                val = value;

        //            if (string.IsNullOrWhiteSpace(val) && (i != (length - 1)))
        //                done = true;
        //            else if (string.IsNullOrWhiteSpace(val) && (i == (length - 1)) && DoubleCheckLR(LR, ws, col))
        //                return (LR-1);
        //            else
        //                LR += incrementsL[i];
                    
        //        } while (!done);

        //        done = false;
        //        LR -= incrementsL[i];
        //        LR += incrementsL[(i + 1)];
        //    }
        //    return 0;
        //}
        private static bool DoubleCheckLR(int LR, Worksheet ws, int col)
        {
            string val = null;
            int qRowsChecked = 25;

            for (int i = 0; i < qRowsChecked; i++)
            {
                var value = ws.Cells[LR, col].Value2;

                if (value is double)
                    val = Convert.ToString(value);
                else
                    val = value;

                if (!string.IsNullOrWhiteSpace(val))
                    return false;

                LR++;
            }
            return true;
        }
        public static int LastRow(WS ws, int firstCol)
        {
            List<int> searchIncrementsList = new List<int>() { 10000, 1000, 100, 1 };
            bool foundEmptyRow = false;
            int searchRow = GetFirstDataRow(ws, firstCol);
            int numberOfRowsToSearchAfterFirstNullValue = 25;
            int lastCol = LastCol(ws, 3);

            for (int i = 0; i < searchIncrementsList.Count; i++)
            {
                do
                {
                    if (!isRowEmpty(ws.Cells[searchRow, firstCol].Value2)) { searchRow += searchIncrementsList[i]; }
                    else if (i == (searchIncrementsList.Count - 1) && !isThereDataAfterNullValue()) 
                        { return (searchRow - 1); }
                    else { foundEmptyRow = true; }
                } while (!foundEmptyRow);

                if(searchRow > searchIncrementsList[i]) { searchRow -= searchIncrementsList[i]; }
                foundEmptyRow = false;
            }
            return 1;

            bool isRowEmpty(object nextCell) => (nextCell == null) ? true : false;
            bool isThereDataAfterNullValue()
            {
                int firstSearchRow = searchRow + numberOfRowsToSearchAfterFirstNullValue;
                for (int row = firstSearchRow; row > searchRow; row--)
                {
                    if (!isRowEmpty(ws.Cells[row, firstCol].Value2)) { return true; }
                }
                return false;
            }
        }
        private static int GetFirstDataRow(WS ws, int col)
        {
            string value;
            int FR = 1;
            do
            {
                value = Convert.ToString(ws.Cells[FR, col].Value2);
                if (!String.IsNullOrEmpty(value)) 
                { 
                    return FR + 1; // assumes first non-null value is header 
                }
                else
                {
                    FR++;
                }

            } while (true);
            return 2;
        }
        private static int GetHeadingsRow(WS ws)
        {
            bool empty, empty2;
            object o;
            int FR = 1;
            do
            {
                o = ws.Cells[FR, 1].Value2;
                empty = (String.IsNullOrEmpty(Convert.ToString(o)));
                empty2 = String.IsNullOrWhiteSpace(Convert.ToString(o));
                FR++;
            } while (empty);
            return FR;
        }
        public static int LastCol(Worksheet ws, int row)
        {
            bool foundEmptyCol = false;
            int searchCol = 1;

            if(ws.Name== "ExpediteReport") { row = 2; }
            else { row = 1; }

            do
            {
                if (ws.Cells[row, searchCol].Value2 == null) { return (searchCol - 1); }
                searchCol++;
            } while (!foundEmptyCol);
            return 1;
        }                
        public static string[,] LoadDirtyArr(Worksheet ws, int lr, int lc)
        {
            string[,] loadArr = new string[lr, lc];
            string val;

            for (int r = 1; r < lr; r++)
            {
                for (int c = 1; c < lc; c++)
                {
                    var v = ws.Cells[r, c].Value2;

                    if (v is string)
                        val = v;
                    else
                        val = Convert.ToString(v);    
                    
                    if (string.IsNullOrWhiteSpace(val?.Trim('"')))
                        val = null;

                    loadArr[r, c] = val;
                }
            }
            return loadArr;
        }
        public static void NewSheetAndWriteCleanArr(string[,] dataArr, Worksheet ws)
        {
            ws.Activate();

            WS newWorksheet = (WS)Globals.ThisAddIn.Application.Worksheets.Add();

            int totRows = dataArr.GetLength(0);
            int totCols = dataArr.GetLength(1);
            bool blankRow = true;

            for (int r = 0; r < totRows; r++)
            {
                for (int i = 0; i < totCols; i++)
                {
                    if (dataArr[r, i] != null)
                        blankRow = false;
                    else
                        blankRow = true;                   
                }
                if (!blankRow)
                {
                    for (int c = 0; c < totCols; c++)
                    {
                        newWorksheet.Cells[(r + 1), (c + 1)].Value2 = dataArr[r, c];
                    }
                }
            }
            newWorksheet.Columns.EntireColumn.AutoFit();
        }        
        public static void FormatAsTable(string TableName)
        {
            var xlApp = new Microsoft.Office.Interop.Excel.Application();

            WB wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            WS ws = wb.ActiveSheet;

            int LR = LastRow(ws,1);
            int LC = LastCol(ws,1);

            Range startCell = ws.Cells[1, 1];
            Range endCell = ws.Cells[LR, LC];

            Range rg = ws.Range[startCell,endCell];

            ws.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, rg, ws, XlYesNoGuess.xlYes, System.Type.Missing).Name = TableName;
            rg.Select();
            // with the below line of code in, in debug mode it errors out...but in live application it works - but seems to stop the code?
            //rg.Worksheet.ListObjects[TableName].TableStyle = "Blue, Table Style Medium 2";
        }
        public static string Quarter(DateTime date)
        {
            int q = 0;

            if (Convert.ToInt32(date.Month) <= 3)
                q = 1;
            else if (Convert.ToInt32(date.Month) <= 6)
                q = 2;
            else if (Convert.ToInt32(date.Month) <= 9)
                q = 3;
            else
                q = 4;

            return (date.Year.ToString() + "-Q" + q);
        }
        //public System.Data.DataTable ConvertToDataTable<T>(IList<T> data)
        public static void NewSheetAndWriteData<T>(IList<T> data, Worksheet ws)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));

            SD.DataTable dt = new SD.DataTable();            

            foreach (PropertyDescriptor prop in properties)
            {
                dt.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            int rowCount = 0;

            foreach (T item in data)
            {
                SD.DataRow row = dt.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                rowCount++;
            }

            int rows = dt.Rows.Count;

            // write data:
            ws.Activate();
            WS newWorksheet = (WS)Globals.ThisAddIn.Application.Worksheets.Add();

            // column headings
            for (var i = 0; i < dt.Columns.Count; i++)
            {
                newWorksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            }

            for (var i = 1; i < rowCount; i++)
            {
                // to do: format datetime values before printing
                for (var j = 1; j < dt.Columns.Count; j++)
                {
                    newWorksheet.Cells[(i + 1), j] = dt.Rows[i]; // [j];
                }
            }
        }
        public static DateTime ReadDateTime(object oDate)
        {
            DateTime dt;
            if(oDate != null && oDate is DateTime)
            {
                dt = (DateTime)oDate;
                return dt.Date;
            }
            else if(oDate is double)
            {
                double d = (double)oDate;
                return DateTime.FromOADate(d);
            }
            else
            {
                return DateTime.MinValue;
            }
        }        
        public static void IfError(KAXLApp kaxlApp)
        {
            kaxlApp.KAXL_RG = new KAXLApp.KAXLRange(kaxlApp, RangeType.Selected);
            var mc = kaxlApp.KAXL_RG;
                                    
            string OriginalFormula;
            string formulaWOEqualSign;
            string modifiedFormula;

            var startCell = kaxlApp.WS.Cells[mc.Row.Start, mc.Col.Start];
            var endCell = kaxlApp.WS.Cells[mc.Row.End, mc.Col.End];
            RG rg = kaxlApp.RG[startCell, endCell];

            for (int row = mc.Row.Start; row <= mc.Row.End; row++)
            {
                for (int col = mc.Col.Start; col <= mc.Col.End; col++)
                {
                    modifiedFormula = null;
                    OriginalFormula = null;
                    formulaWOEqualSign = null;

                    OriginalFormula = kaxlApp.WS.Cells[row,col].Formula;
                    formulaWOEqualSign = OriginalFormula.Substring(1); // removes "=" sign form original formula
                    modifiedFormula = "=iferror((" + formulaWOEqualSign + "),0)";
                    kaxlApp.WS.Cells[row, col].Formula = modifiedFormula;
                }
            }                       
        }
        public static void OverWriteFormulas(KAXLApp kaxlApp)
        {
            kaxlApp.KAXL_RG = new KAXLApp.KAXLRange(kaxlApp, RangeType.Selected);
            var mc = kaxlApp.KAXL_RG;
            int LastRow = KAXL.LastRow(kaxlApp.WS, 1);
            var ws = kaxlApp.WS;

            for (int col = mc.Col.Start; col <= mc.Col.End; col++)
            {
                var startCell = ws.Cells[mc.Row.Start, col];
                var endCell = ws.Cells[LastRow, col];

                RG rg = ws.Range[startCell, endCell];                
                var val = rg.Value2;
                rg.Value2 = val;            
            }
        }
        public static void TopRowFormulas(KAXLApp kaxlApp)
        {
            Cursor c = Cursors.WaitCursor;

            kaxlApp.KAXL_RG = new KAXLApp.KAXLRange(kaxlApp, RangeType.Selected);
            var mc = kaxlApp.KAXL_RG;
            int LastRow = KAXL.LastRow(kaxlApp.WS, 1);
            var ws = kaxlApp.WS;
            int C = 0;

            try
            {
                for (int col = mc.Col.Start; col <= mc.Col.End; col++)
                {
                    C = col;
                    string extractedFormula = ws.Cells[1, col].Value2;

                    var startCell = ws.Cells[mc.Row.Start, col];
                    var endCell = ws.Cells[LastRow, col];

                    RG rg = ws.Range[startCell, endCell];
                    rg.Formula = extractedFormula;
                    var val = rg.Value2;
                    rg.Value2 = val;
                }
            }
            catch
            {
                MessageBox.Show("F'ed up.....formula is f'ed up for column " + ws.Cells[2, C].Value2 + ".  Fix it Jackass");
            }
            c = Cursors.Default;
        }
        public static void CADtoUSDConverter(KAXLApp kaxlApp, double exRate)
        {
            kaxlApp.KAXL_RG = new KAXLApp.KAXLRange(kaxlApp, RangeType.Selected);
            var mc = kaxlApp.KAXL_RG;            
            var ws = kaxlApp.WS;

            for (int iRow = mc.Row.Start; iRow < mc.Row.End; iRow++)
            {
                var CAD = ws.Cells[iRow, mc.Col.Start].Value2;
                if (CAD is string)
                {
                    if(double.TryParse(CAD, out double r))
                        ws.Cells[iRow, mc.Col.Start].Value2 = r;
                }
                else if(CAD is double)
                {
                    ws.Cells[iRow, mc.Col.Start].Value2 = CAD * exRate;
                }
            }
        }
        public static void ScrubItemNumbers(KAXLApp kaxlApp)
        {
            kaxlApp.KAXL_RG = new KAXLApp.KAXLRange(kaxlApp, RangeType.Selected);
            var mc = kaxlApp.KAXL_RG;
            var ws = kaxlApp.WS;

            for (int iRow = mc.Row.Start; iRow < mc.Row.End; iRow++)
            {
                var val = ws.Cells[iRow, mc.Col.Start].Value2;
                if(double.TryParse(val,out double result))
                {
                    ws.Cells[iRow, mc.Col.Start].Value2 = result;
                }
                else
                {
                    ws.Cells[iRow, mc.Col.Start].Value2 = val;
                }
            }
        }
        public static void ExportAsPDF(xlApp xlapp)
        {
            

            //xlWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, 
              //  filepath, Excel.XlFixedFormatQuality.xlQualityStandard, true, false, 1, 1, true, missing);
        }
        public static int FindFirstRowAfterHeader(WS ws) => ws.Cells[1, 1] == null ? 3 : 2;
    }
    public class ColIDL
    {
        private List<string> columnHeadingsL;

        public ColIDL() { }

        public ColIDL(WS ws)
        {
            columnHeadingsL = new List<string>() { null };
            int LC = KAXL.LastCol(ws, 1);
            int tableStartRow = 0;
            bool isStartRow = false;

            do
            {
                tableStartRow++;
                isStartRow = ws.Cells[tableStartRow, 1].Value2 == null ? false : true;
            } while (!isStartRow);

            for (int i = 1; i <= LC; i++)
            {
                columnHeadingsL.Add(ws.Cells[tableStartRow, i].Value2);
            }
        }
        public ColIDL(List<string> _listOfColumnHeadings) => columnHeadingsL = _listOfColumnHeadings; 
        public int GetColNum(string heading)
        {
            int colNum = columnHeadingsL.IndexOf(heading);

            if(colNum != (-1))
            {
                return colNum;
            }
            else
            {
                MessageBox.Show("The source data is missing the " + heading + " field.");
                return colNum;
            }
        }
        public int ColQ => columnHeadingsL.Count;
    }
    public enum RangeType { Selected, WorkSheet, CodedRangeSetKAXLAppRG } // set this up as optional parameter for constructor?
    public class KAXLApp
    {
        public xlApp XL { get; set; }
        public WB WB { get; set; }
        public WS WS { get; set; }
        public RG RG { get; set; }
        public KAXLRange KAXL_RG { get; set; }
        public Process Process { get; set; }
        public KAXLErrorTracker ErrorTracker { get; set; }

        public KAXLApp()
        {
            XL = new Microsoft.Office.Interop.Excel.Application();
            WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            WS = WB.ActiveSheet;
            RG = Globals.ThisAddIn.Application.Selection;
            KAXL_RG = new KAXLRange();
            ErrorTracker = new KAXLErrorTracker();
        }

        public KAXLApp(string pathToFileToOpen, int sheetNumberToOpen = 1)
        {
            XL = new Microsoft.Office.Interop.Excel.Application();

            WB = XL.Workbooks.Open(pathToFileToOpen);
            WS = WB.Sheets[sheetNumberToOpen];
            KAXL_RG = new KAXLRange(this, RangeType.WorkSheet);
            ErrorTracker = new KAXLErrorTracker();
        }

        public static void CloseSheet(KAXLApp xlapp)
        {
            if(xlapp.Process != null)
            {
                try
                {
                    xlapp.Process.Kill();
                    xlapp.Process.Dispose();
                }
                catch
                {
                    MessageBox.Show("Something's f'd up...");
                }
            }
            else
            {
                xlapp.WB.Save();
                xlapp.WB.Close();
                xlapp.XL.Quit();
            }
            ReleaseObject(xlapp.RG);
            ReleaseObject(xlapp.WS);
            ReleaseObject(xlapp.WB);
            ReleaseObject(xlapp.XL);
            ReleaseObject(xlapp.Process);
            ReleaseObject(xlapp);
        }
        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        public class KAXLRange
        {
            public int Start { get; set; }
            public int End { get; set; }
            public int Q { get; set; }
            public int Max_Array { get; set; }
            public Row Row { get; set; }
            public Col Col { get; set; }
            public RG RG { get; set; }
            
            public KAXLRange() { }            

            public KAXLRange(KAXLApp kaxlApp, RangeType rt)
            {                
                if(rt == RangeType.Selected)
                {
                    Row = new Row();
                    Row.Start = kaxlApp.RG.Row;
                    Row.Q = kaxlApp.RG.Rows.Count;
                    Row.End = Row.Start + Row.Q;
                    
                    Col = new Col();
                    Col.Start = kaxlApp.RG.Column;
                    Col.Q = kaxlApp.RG.Columns.Count;
                    Col.End = Col.Start + Col.Q;

                    RG = kaxlApp.WS.Range[kaxlApp.WS.Cells[Row.Start, Col.Start], kaxlApp.WS.Cells[Row.End, Col.End]];
                }
                else if(rt == RangeType.WorkSheet)
                {
                    int FR = KAXL.FindFirstRowAfterHeader(kaxlApp.WS);
                    int LR = KAXL.LastRow(kaxlApp.WS, 1);
                    object test = kaxlApp.WS.Cells[3, 1].Value2;

                    Row = new Row()
                    {
                        Start = KAXL.FindFirstRowAfterHeader(kaxlApp.WS),
                        End = KAXL.LastRow(kaxlApp.WS, 1),
                    };
                    Row.Q = Row.End - Row.Start + 1;

                    Col = new Col()
                    {
                        Start = 1,
                        End = KAXL.LastCol(kaxlApp.WS, Row.Start),
                    };
                    Col.Q = Col.End - Col.Start + 1;

                    Row.Max_Array = Row.Q + 1;
                    Col.Max_Array = Col.Q + 1;

                    RG = kaxlApp.WS.Range[kaxlApp.WS.Cells[Row.Start, Col.Start], kaxlApp.WS.Cells[Row.End, Col.End]];
                }
                else if(rt == RangeType.CodedRangeSetKAXLAppRG)
                {
                    Row = new Row()
                    {
                        Start = kaxlApp.RG.Row,
                        Q = kaxlApp.RG.Rows.Count                      
                    };
                    Row.End = Row.Start + Row.Q - 1;

                    Col = new Col()
                    {
                        Start = kaxlApp.RG.Column,
                        Q = kaxlApp.RG.Columns.Count
                    };
                    Col.End = Col.Start + Col.Q;
                        
                    RG = kaxlApp.RG;
                }
                _valueArray = (object[,])RG.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            }

            private readonly object[,] _valueArray;

            public object this[int r, int c]
            {
                get => _valueArray[r, c];
                set => _valueArray[r, c] = value;
            }

        }
        public class Row : KAXLRange { }
        public class Col : KAXLRange { }
    }
    public class KAXLErrorTracker
    {
        private List<string> _errorList;

        public KAXLErrorTracker()
        {
            _errorList = new List<string>();
        }

        public int Row { get; set; }        
        public string ProgramStage { get; set; }
        public void AddNewError(string error) => _errorList.Add(error);

        public string this[int i]
        {
            get => _errorList[i];
            set => _errorList[i] = value;
        }
        public List<string> GetErrorList() => _errorList;
    }
}
