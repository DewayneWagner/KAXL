using System;
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
using Microsoft.Office.Tools.Excel;
using System.Collections;

namespace DKAExcelStuff
{
    // This library will contain various code and methods that will be used frequently when using VSTO

    public class KAXL
    {   
        public static int LastRow(Worksheet ws, int col)
        {
            string val = null;
            bool done = false;

            List<int> incrementsL = new List<int>() { 10000, 1000, 100, 1 };
            int length = incrementsL.Count;

            int LR = incrementsL[0] + 1;

            for (int i = 0; i < length; i++)
            {
                do
                {
                    var value = ws.Cells[LR, col].Value2;

                    if (value is double)
                        val = Convert.ToString(value);
                    else
                        val = value;

                    if (string.IsNullOrWhiteSpace(val) && (i != (length - 1)))
                        done = true;
                    else if (string.IsNullOrWhiteSpace(val) && (i == (length - 1)) && DoubleCheckLR(LR, ws, col))
                        return (LR-1);
                    else
                        LR += incrementsL[i];
                    
                } while (!done);

                done = false;
                LR -= incrementsL[i];
                LR += incrementsL[(i + 1)];
            }
            return 0;
        }
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
        public static int LastCol(Worksheet ws, int row)
        {
            string val = null;

            for (int LC = 50; LC >= 1; LC--)
            {
                var value = ws.Cells[row, LC].Value2;

                if (value is double)
                    val = Convert.ToString(value);
                else
                    val = value;

                if (val != null)
                    return (LC);
            }
            return 0;
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
        public static string Quarter(int month, int year)
        {
            int q = 0;

            if (month <= 3)
                q = 1;
            else if (month <= 6)
                q = 2;
            else if (month <= 9)
                q = 3;
            else
                q = 4;

            return (year + "-Q" + q);
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
        public static int YearFromString(string date)
        {
            DateTime d = ReadDateTime(date);
            return d.Year;
        }
        public static int MonthFromString(string date)
        {
            DateTime d = ReadDateTime(date);
            return d.Month;
        }
        public static DateTime ReadDateTime(object oDate)
        {
            try
            {
                double d;
                DateTime dt;

                if (oDate == null)
                {
                    dt = DateTime.MinValue;
                }
                else if (oDate is string)
                {
                    dt = Convert.ToDateTime(oDate);
                    dt = DateTime.FromOADate(Math.Floor(dt.ToOADate()));
                }
                else if (oDate is DateTime)
                {
                    d = Math.Round(Convert.ToDouble(oDate), 0);
                    dt = DateTime.FromOADate(d);
                }
                else
                {
                    dt = DateTime.MinValue;
                }
                dt.ToShortDateString();
                return dt;
            }
            catch
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
            kaxlApp.KAXL_RG = new KAXLApp.KAXLRange(kaxlApp, RangeType.Selected);
            var mc = kaxlApp.KAXL_RG;
            int LastRow = KAXL.LastRow(kaxlApp.WS, 1);
            var ws = kaxlApp.WS;            

            for (int col = mc.Col.Start; col <= mc.Col.End; col++)
            {
                string extractedFormula = ws.Cells[1, col].Value2;

                var startCell = ws.Cells[mc.Row.Start, col];
                var endCell = ws.Cells[LastRow, col];

                RG rg = ws.Range[startCell, endCell];
                rg.Formula = extractedFormula;
                var val = rg.Value2;
                rg.Value2 = val;
            }
        }
        public static void CADtoUSDConverter(KAXLApp kaxlApp, double exRate)
        {
            kaxlApp.KAXL_RG = new KAXLApp.KAXLRange(kaxlApp, RangeType.Selected);
            var mc = kaxlApp.KAXL_RG;            
            var ws = kaxlApp.WS;
            //double exRate = 0.76037; //02-28-2019
            
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
            int LC = KAXL.LastCol(ws,1);
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

        public KAXLApp()
        {
            XL = new Microsoft.Office.Interop.Excel.Application();
            WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            WS = WB.ActiveSheet;
            RG = Globals.ThisAddIn.Application.Selection;
            KAXL_RG = new KAXLRange();
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
            public NR NamedRange { get; set; }

            public KAXLRange() { }            

            public KAXLRange(KAXLApp kaxlApp, RangeType rt)
            {
                NamedRange = new NR();
                if(rt == RangeType.Selected)
                {                    
                    Row = new Row()
                    {
                        Start = kaxlApp.RG.Row,
                        Q = kaxlApp.RG.Rows.Count,
                        End = Start + Q,
                    };

                    Col = new Col()
                    {
                        Start = kaxlApp.RG.Column,
                        Q = kaxlApp.RG.Columns.Count,
                        End = Start + Q,
                    };

                    RG = kaxlApp.WS.Range[kaxlApp.WS.Cells[Row.Start, Col.Start], kaxlApp.WS.Cells[Row.End, Col.End]];
                }
                else if(rt == RangeType.WorkSheet)
                {
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
    public class NR : NamedRange
    {
        public void Delete()
        {
            throw new NotImplementedException();
        }

        public object get_Value(object rangeValueDataType)
        {
            throw new NotImplementedException();
        }

        public void set_Value(object rangeValueDataType, object arg1)
        {
            throw new NotImplementedException();
        }

        public IEnumerator GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public object Activate()
        {
            throw new NotImplementedException();
        }

        public string get_Address(object RowAbsolute, object ColumnAbsolute, XlReferenceStyle ReferenceStyle = XlReferenceStyle.xlA1, object External = null, object RelativeTo = null)
        {
            throw new NotImplementedException();
        }

        public string get_AddressLocal(object RowAbsolute, object ColumnAbsolute, XlReferenceStyle ReferenceStyle = XlReferenceStyle.xlA1, object External = null, object RelativeTo = null)
        {
            throw new NotImplementedException();
        }

        public object AdvancedFilter(XlFilterAction Action, object CriteriaRange, object CopyToRange, object Unique)
        {
            throw new NotImplementedException();
        }

        public object ApplyNames(object Names, object IgnoreRelativeAbsolute, object UseRowColumnNames, object OmitColumn, object OmitRow, XlApplyNamesOrder Order = XlApplyNamesOrder.xlRowThenColumn, object AppendLast = null)
        {
            throw new NotImplementedException();
        }

        public object ApplyOutlineStyles()
        {
            throw new NotImplementedException();
        }

        public string AutoComplete(string String)
        {
            throw new NotImplementedException();
        }

        public object AutoFill(RG Destination, XlAutoFillType Type = XlAutoFillType.xlFillDefault)
        {
            throw new NotImplementedException();
        }

        public object AutoFilter(object Field, object Criteria1, XlAutoFilterOperator Operator = XlAutoFilterOperator.xlAnd, object Criteria2 = null, object VisibleDropDown = null)
        {
            throw new NotImplementedException();
        }

        public object AutoFit()
        {
            throw new NotImplementedException();
        }

        public object AutoFormat(XlRangeAutoFormat Format = XlRangeAutoFormat.xlRangeAutoFormatClassic1, object Number = null, object Font = null, object Alignment = null, object Border = null, object Pattern = null, object Width = null)
        {
            throw new NotImplementedException();
        }

        public object AutoOutline()
        {
            throw new NotImplementedException();
        }

        public object BorderAround(object LineStyle, XlBorderWeight Weight = XlBorderWeight.xlThin, XlColorIndex ColorIndex = XlColorIndex.xlColorIndexAutomatic, object Color = null)
        {
            throw new NotImplementedException();
        }

        public object Calculate()
        {
            throw new NotImplementedException();
        }

        public Characters get_Characters(object Start, object Length)
        {
            throw new NotImplementedException();
        }

        public object CheckSpelling(object CustomDictionary, object IgnoreUppercase, object AlwaysSuggest, object SpellLang)
        {
            throw new NotImplementedException();
        }

        public object Clear()
        {
            throw new NotImplementedException();
        }

        public object ClearContents()
        {
            throw new NotImplementedException();
        }

        public object ClearFormats()
        {
            throw new NotImplementedException();
        }

        public object ClearNotes()
        {
            throw new NotImplementedException();
        }

        public object ClearOutline()
        {
            throw new NotImplementedException();
        }

        public RG ColumnDifferences(object Comparison)
        {
            throw new NotImplementedException();
        }

        public object Consolidate(object Sources, object Function, object TopRow, object LeftColumn, object CreateLinks)
        {
            throw new NotImplementedException();
        }

        public object Copy(object Destination)
        {
            throw new NotImplementedException();
        }

        public int CopyFromRecordset(object Data, object MaxRows, object MaxColumns)
        {
            throw new NotImplementedException();
        }

        public object CopyPicture(XlPictureAppearance Appearance = XlPictureAppearance.xlScreen, XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture)
        {
            throw new NotImplementedException();
        }

        public object CreateNames(object Top, object Left, object Bottom, object Right)
        {
            throw new NotImplementedException();
        }

        public object CreatePublisher(object Edition, XlPictureAppearance Appearance = XlPictureAppearance.xlScreen, object ContainsPICT = null, object ContainsBIFF = null, object ContainsRTF = null, object ContainsVALU = null)
        {
            throw new NotImplementedException();
        }

        public object Cut(object Destination)
        {
            throw new NotImplementedException();
        }

        public object DataSeries(object Rowcol, XlDataSeriesType Type = XlDataSeriesType.xlDataSeriesLinear, XlDataSeriesDate Date = XlDataSeriesDate.xlDay, object Step = null, object Stop = null, object Trend = null)
        {
            throw new NotImplementedException();
        }

        public object DialogBox()
        {
            throw new NotImplementedException();
        }

        public object EditionOptions(XlEditionType Type, XlEditionOptionsOption Option, object Name, object Reference, XlPictureAppearance Appearance = XlPictureAppearance.xlScreen, XlPictureAppearance ChartSize = XlPictureAppearance.xlScreen, object Format = null)
        {
            throw new NotImplementedException();
        }

        public RG get_End(XlDirection Direction)
        {
            throw new NotImplementedException();
        }

        public object FillDown()
        {
            throw new NotImplementedException();
        }

        public object FillLeft()
        {
            throw new NotImplementedException();
        }

        public object FillRight()
        {
            throw new NotImplementedException();
        }

        public object FillUp()
        {
            throw new NotImplementedException();
        }

        public RG Find(object What, object After, object LookIn, object LookAt, object SearchOrder, XlSearchDirection SearchDirection = XlSearchDirection.xlNext, object MatchCase = null, object MatchByte = null, object SearchFormat = null)
        {
            throw new NotImplementedException();
        }

        public RG FindNext(object After)
        {
            throw new NotImplementedException();
        }

        public RG FindPrevious(object After)
        {
            throw new NotImplementedException();
        }

        public object FunctionWizard()
        {
            throw new NotImplementedException();
        }

        public bool GoalSeek(object Goal, RG ChangingCell)
        {
            throw new NotImplementedException();
        }

        public object Group(object Start, object End, object By, object Periods)
        {
            throw new NotImplementedException();
        }

        public void InsertIndent(int InsertAmount)
        {
            throw new NotImplementedException();
        }

        public object Insert(object Shift, object CopyOrigin)
        {
            throw new NotImplementedException();
        }

        public object get_Item(object RowIndex, object ColumnIndex)
        {
            throw new NotImplementedException();
        }

        public void set_Item(object RowIndex, object ColumnIndex, object _arg2)
        {
            throw new NotImplementedException();
        }

        public object Justify()
        {
            throw new NotImplementedException();
        }

        public object ListNames()
        {
            throw new NotImplementedException();
        }

        public void Merge(object Across)
        {
            throw new NotImplementedException();
        }

        public void UnMerge()
        {
            throw new NotImplementedException();
        }

        public object NavigateArrow(object TowardPrecedent, object ArrowNumber, object LinkNumber)
        {
            throw new NotImplementedException();
        }

        public string NoteText(object Text, object Start, object Length)
        {
            throw new NotImplementedException();
        }

        public RG get_Offset(object RowOffset, object ColumnOffset)
        {
            throw new NotImplementedException();
        }

        public object Parse(object ParseLine, object Destination)
        {
            throw new NotImplementedException();
        }

        public object _PasteSpecial(XlPasteType Paste = XlPasteType.xlPasteAll, XlPasteSpecialOperation Operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone, object SkipBlanks = null, object Transpose = null)
        {
            throw new NotImplementedException();
        }

        public object _PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate)
        {
            throw new NotImplementedException();
        }

        public object PrintPreview(object EnableChanges)
        {
            throw new NotImplementedException();
        }

        public RG get_Range(object Cell1, object Cell2)
        {
            throw new NotImplementedException();
        }

        public object RemoveSubtotal()
        {
            throw new NotImplementedException();
        }

        public bool Replace(object What, object Replacement, object LookAt, object SearchOrder, object MatchCase, object MatchByte, object SearchFormat, object ReplaceFormat)
        {
            throw new NotImplementedException();
        }

        public RG get_Resize(object RowSize, object ColumnSize)
        {
            throw new NotImplementedException();
        }

        public RG RowDifferences(object Comparison)
        {
            throw new NotImplementedException();
        }

        public object Run(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            throw new NotImplementedException();
        }

        public object Select()
        {
            throw new NotImplementedException();
        }

        public object Show()
        {
            throw new NotImplementedException();
        }

        public object ShowDependents(object Remove)
        {
            throw new NotImplementedException();
        }

        public object ShowErrors()
        {
            throw new NotImplementedException();
        }

        public object ShowPrecedents(object Remove)
        {
            throw new NotImplementedException();
        }

        public object Sort(object Key1, XlSortOrder Order1 = XlSortOrder.xlAscending, object Key2 = null, object Type = null, XlSortOrder Order2 = XlSortOrder.xlAscending, object Key3 = null, XlSortOrder Order3 = XlSortOrder.xlAscending, XlYesNoGuess Header = XlYesNoGuess.xlNo, object OrderCustom = null, object MatchCase = null, XlSortOrientation Orientation = XlSortOrientation.xlSortRows, XlSortMethod SortMethod = XlSortMethod.xlPinYin, XlSortDataOption DataOption1 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption2 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption3 = XlSortDataOption.xlSortNormal)
        {
            throw new NotImplementedException();
        }

        public object SortSpecial(XlSortMethod SortMethod = XlSortMethod.xlPinYin, object Key1 = null, XlSortOrder Order1 = XlSortOrder.xlAscending, object Type = null, object Key2 = null, XlSortOrder Order2 = XlSortOrder.xlAscending, object Key3 = null, XlSortOrder Order3 = XlSortOrder.xlAscending, XlYesNoGuess Header = XlYesNoGuess.xlNo, object OrderCustom = null, object MatchCase = null, XlSortOrientation Orientation = XlSortOrientation.xlSortRows, XlSortDataOption DataOption1 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption2 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption3 = XlSortDataOption.xlSortNormal)
        {
            throw new NotImplementedException();
        }

        public RG SpecialCells(XlCellType Type, object Value)
        {
            throw new NotImplementedException();
        }

        public object SubscribeTo(string Edition, XlSubscribeToFormat Format = XlSubscribeToFormat.xlSubscribeToText)
        {
            throw new NotImplementedException();
        }

        public object Subtotal(int GroupBy, XlConsolidationFunction Function, object TotalList, object Replace, object PageBreaks, XlSummaryRow SummaryBelowData = XlSummaryRow.xlSummaryBelow)
        {
            throw new NotImplementedException();
        }

        public object Table(object RowInput, object ColumnInput)
        {
            throw new NotImplementedException();
        }

        public object TextToColumns(object Destination, XlTextParsingType DataType = XlTextParsingType.xlDelimited, XlTextQualifier TextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote, object ConsecutiveDelimiter = null, object Tab = null, object Semicolon = null, object Comma = null, object Space = null, object Other = null, object OtherChar = null, object FieldInfo = null, object DecimalSeparator = null, object ThousandsSeparator = null, object TrailingMinusNumbers = null)
        {
            throw new NotImplementedException();
        }

        public object Ungroup()
        {
            throw new NotImplementedException();
        }

        public Comment AddComment(object Text)
        {
            throw new NotImplementedException();
        }

        public void ClearComments()
        {
            throw new NotImplementedException();
        }

        public void SetPhonetic()
        {
            throw new NotImplementedException();
        }

        public object PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }

        public void Dirty()
        {
            throw new NotImplementedException();
        }

        public void Speak(object SpeakDirection, object SpeakFormulas)
        {
            throw new NotImplementedException();
        }

        public object PasteSpecial(XlPasteType Paste = XlPasteType.xlPasteAll, XlPasteSpecialOperation Operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone, object SkipBlanks = null, object Transpose = null)
        {
            throw new NotImplementedException();
        }

        public void RemoveDuplicates(object Columns, XlYesNoGuess Header = XlYesNoGuess.xlNo)
        {
            throw new NotImplementedException();
        }

        public object PrintOutEx(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }

        public void ExportAsFixedFormat(XlFixedFormatType Type, object Filename, object Quality, object IncludeDocProperties, object IgnorePrintAreas, object From, object To, object OpenAfterPublish, object FixedFormatExtClassPtr)
        {
            throw new NotImplementedException();
        }

        public object CalculateRowMajorOrder()
        {
            throw new NotImplementedException();
        }

        public void ClearHyperlinks()
        {
            throw new NotImplementedException();
        }

        public object BorderAround2(object LineStyle, XlBorderWeight Weight = XlBorderWeight.xlThin, XlColorIndex ColorIndex = XlColorIndex.xlColorIndexAutomatic, object Color = null, object ThemeColor = null)
        {
            throw new NotImplementedException();
        }

        public void AllocateChanges()
        {
            throw new NotImplementedException();
        }

        public void DiscardChanges()
        {
            throw new NotImplementedException();
        }

        public IContainer Container => throw new NotImplementedException();

        public object Tag { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public NamedRange_AddressType Address => throw new NotImplementedException();

        public NamedRange_AddressLocalType AddressLocal => throw new NotImplementedException();

        public NamedRange_CharactersType Characters => throw new NotImplementedException();

        public NamedRange_ItemType Item => throw new NotImplementedException();

        public NamedRange_OffsetType Offset => throw new NotImplementedException();

        public NamedRange_ResizeType Resize => throw new NotImplementedException();

        public RG InnerObject => throw new NotImplementedException();

        public string RefersTo { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public string RefersToLocal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public string RefersToR1C1 { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public string RefersToR1C1Local { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public RG RefersToRange => throw new NotImplementedException();

        public DataSourceUpdateMode DefaultDataSourceUpdateMode { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object Value { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object Name { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public xlApp Application => throw new NotImplementedException();

        public XlCreator Creator => throw new NotImplementedException();

        public object Parent => throw new NotImplementedException();

        public object AddIndent { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Areas Areas => throw new NotImplementedException();

        public Borders Borders => throw new NotImplementedException();

        public RG Cells => throw new NotImplementedException();

        public int Column => throw new NotImplementedException();

        public RG Columns => throw new NotImplementedException();

        public object ColumnWidth { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Count => throw new NotImplementedException();

        public RG CurrentArray => throw new NotImplementedException();

        public RG CurrentRegion => throw new NotImplementedException();

        public RG Dependents => throw new NotImplementedException();

        public RG DirectDependents => throw new NotImplementedException();

        public RG DirectPrecedents => throw new NotImplementedException();

        public NamedRange_EndType End => throw new NotImplementedException();

        public RG EntireColumn => throw new NotImplementedException();

        public RG EntireRow => throw new NotImplementedException();

        public Font Font => throw new NotImplementedException();

        public object Formula { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object FormulaArray { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public XlFormulaLabel FormulaLabel { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object FormulaHidden { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object FormulaLocal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object FormulaR1C1 { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object FormulaR1C1Local { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public object HasArray => throw new NotImplementedException();

        public object HasFormula => throw new NotImplementedException();

        public object Height => throw new NotImplementedException();

        public object Hidden { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object HorizontalAlignment { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object IndentLevel { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Interior Interior => throw new NotImplementedException();

        public object Left => throw new NotImplementedException();

        public int ListHeaderRows => throw new NotImplementedException();

        public XlLocationInTable LocationInTable => throw new NotImplementedException();

        public object Locked { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public RG MergeArea => throw new NotImplementedException();

        public object MergeCells { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public RG Next => throw new NotImplementedException();

        public object NumberFormat { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object NumberFormatLocal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object Orientation { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object OutlineLevel { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public int PageBreak { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public PivotField PivotField => throw new NotImplementedException();

        public PivotItem PivotItem => throw new NotImplementedException();

        public PivotTable PivotTable => throw new NotImplementedException();

        public RG Precedents => throw new NotImplementedException();

        public object PrefixCharacter => throw new NotImplementedException();

        public RG Previous => throw new NotImplementedException();

        public QueryTable QueryTable => throw new NotImplementedException();

        public int Row => throw new NotImplementedException();

        public object RowHeight { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public RG Rows => throw new NotImplementedException();

        public object ShowDetail { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object ShrinkToFit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public SoundNote SoundNote => throw new NotImplementedException();

        public object Style { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public object Summary => throw new NotImplementedException();

        public object Text => throw new NotImplementedException();

        public object Top => throw new NotImplementedException();

        public object UseStandardHeight { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object UseStandardWidth { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Validation Validation => throw new NotImplementedException();

        public object Value2 { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public object VerticalAlignment { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public object Width => throw new NotImplementedException();

        public WS Worksheet => throw new NotImplementedException();

        public object WrapText { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Comment Comment => throw new NotImplementedException();

        public Phonetic Phonetic => throw new NotImplementedException();

        public FormatConditions FormatConditions => throw new NotImplementedException();

        public int ReadingOrder { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Hyperlinks Hyperlinks => throw new NotImplementedException();

        public Phonetics Phonetics => throw new NotImplementedException();

        public string ID { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public PivotCell PivotCell => throw new NotImplementedException();

        public Errors Errors => throw new NotImplementedException();

        public SmartTags SmartTags => throw new NotImplementedException();

        public bool AllowEdit => throw new NotImplementedException();

        public Microsoft.Office.Interop.Excel.ListObject ListObject => throw new NotImplementedException();

        public XPath XPath => throw new NotImplementedException();

        public Actions ServerActions => throw new NotImplementedException();

        public string MDX => throw new NotImplementedException();

        public object CountLarge => throw new NotImplementedException();

        public SparklineGroups SparklineGroups => throw new NotImplementedException();

        public DisplayFormat DisplayFormat => throw new NotImplementedException();

        public event DocEvents_BeforeDoubleClickEventHandler BeforeDoubleClick;
        public event DocEvents_BeforeRightClickEventHandler BeforeRightClick;
        public event DocEvents_SelectionChangeEventHandler SelectionChange;
        public event DocEvents_SelectionChangeEventHandler Selected;
        public event DocEvents_SelectionChangeEventHandler Deselected;
        public event EventHandler BindingContextChanged;
        public event DocEvents_ChangeEventHandler Change;

        public ControlBindingsCollection DataBindings => throw new NotImplementedException();

        public BindingContext BindingContext { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public ISite Site { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public event EventHandler Disposed;

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public void BeginInit()
        {
            throw new NotImplementedException();
        }

        public void EndInit()
        {
            throw new NotImplementedException();
        }
    }
}
