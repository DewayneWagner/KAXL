using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using Microsoft.Office.Tools.Ribbon;
using NR = Microsoft.Office.Tools.Excel.NamedRange;
using DKARibbon;

namespace DKAExcelStuff
{
    public enum RowTypeE
    {
        MainTitle, Blank, MainTitleBlock, Employee, ReqTitle, RightTitleBlock,
        ReqID, SubHeadings, LineData, SecSubTotals, EmpTotals, GrandTotals, TotCol
    };
    public enum CleanCol
    {
        Employee, ReqID, ProjectName, ProdName,Quantity,UnitPrice,Vendor,NetAmount,Currency,TotCol
    }
    public enum DirtyCol
    {
        EmpName = 2,
        ReqID = 2,
        ProjName = 6,
        ProductName = 6,
        Quantity = 7,
        UnitPrice = 8,
        Vendor = 13,
        NetAmount = 14,
        Currency = 17,
        TotCol =22
    }
    public class WarRoomFormat
    {
        public static void ReqData()
        {
            var xlApp = new Application();

            WB wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            WS ws = wb.ActiveSheet;

            int LR = KAXL.LastRow(ws,2);
            int LC = (int)DirtyCol.TotCol;
            string[,] RowDataArr = new string[LR, LC];

            UnmergeAll(ws, LR, LC);

            int rowType = 0, nextCleanRow = 1;
            
            // build an array of dirty data
            string[,] DirtyDataA = KAXL.LoadDirtyArr(ws, LR, LC);

            // array for scrubbed data
            string[,] CleanDataA = new string[LR, (int)CleanCol.TotCol];

            for (int i = 0; i < (int)CleanCol.TotCol; i++)
            {
                CleanDataA[0, i] = ((CleanCol)i).ToString();
            }

            // list for sending 1 row of data at a time to method to classify row type
            string[] RowDataA = new string[LC];

            // to reuse Employee name and ReqID
            string empName = "", reqID = "", projName = "";

            for (int iRow = 1; iRow < LR; iRow++)
            {
                for (int i = 0; i < LC; i++)
                {
                    RowDataA[i]=(DirtyDataA[iRow, i]);
                }
                rowType = RowType(RowDataA);

                switch (rowType)
                {
                    case (int)RowTypeE.MainTitle:
                    case (int)RowTypeE.Blank:
                    case (int)RowTypeE.MainTitleBlock:
                    case (int)RowTypeE.ReqTitle:
                    case (int)RowTypeE.SubHeadings:
                    case (int)RowTypeE.SecSubTotals:
                    case (int)RowTypeE.EmpTotals:
                    case (int)RowTypeE.GrandTotals:
                    case (int)RowTypeE.RightTitleBlock:
                        break;

                    case (int)RowTypeE.Employee:
                        empName = DirtyDataA[iRow, (int)DirtyCol.EmpName];
                        CleanDataA[nextCleanRow, (int)CleanCol.Employee] = empName;
                        break;
                    case (int)RowTypeE.ReqID:
                        reqID = DirtyDataA[iRow, (int)DirtyCol.ReqID];
                        projName = DirtyDataA[iRow, (int)DirtyCol.ProjName];
                        CleanDataA[nextCleanRow, (int)CleanCol.ReqID] = reqID;
                        CleanDataA[nextCleanRow, (int)CleanCol.ProjectName] = projName;
                        break;
                    case (int)RowTypeE.LineData:

                        if (empName == null)
                            empName = CleanDataA[(nextCleanRow - 1), (int)CleanCol.Employee];
                        else
                            CleanDataA[nextCleanRow, (int)CleanCol.Employee] = empName;

                        CleanDataA[nextCleanRow, (int)CleanCol.ReqID] = reqID;
                        CleanDataA[nextCleanRow, (int)CleanCol.ProjectName] = projName;

                        CleanDataA[nextCleanRow, (int)CleanCol.ProdName] = DirtyDataA[iRow,(int)DirtyCol.ProductName];
                        CleanDataA[nextCleanRow, (int)CleanCol.Quantity] = DirtyDataA[iRow,(int)DirtyCol.Quantity];
                        CleanDataA[nextCleanRow, (int)CleanCol.UnitPrice] = DirtyDataA[iRow, (int)DirtyCol.UnitPrice];
                        CleanDataA[nextCleanRow, (int)CleanCol.Vendor] = DirtyDataA[iRow, (int)DirtyCol.Vendor];
                        CleanDataA[nextCleanRow, (int)CleanCol.NetAmount] = DirtyDataA[iRow, (int)DirtyCol.NetAmount];
                        CleanDataA[nextCleanRow, (int)CleanCol.Currency] = DirtyDataA[iRow, (int)DirtyCol.Currency];
                        nextCleanRow++;
                        iRow++;

                        for (int i = 0; i < (int)DirtyCol.TotCol; i++)
                        {
                            RowDataA[i] = (DirtyDataA[iRow, i]);
                        }
                        rowType = RowType(RowDataA);

                        while (rowType == (int)RowTypeE.LineData)
                        {
                            CleanDataA[nextCleanRow, (int)CleanCol.Employee] = empName;
                            CleanDataA[nextCleanRow, (int)CleanCol.ReqID] = reqID;
                            CleanDataA[nextCleanRow, (int)CleanCol.ProjectName] = projName;

                            CleanDataA[nextCleanRow, (int)CleanCol.ProdName] = DirtyDataA[iRow, (int)DirtyCol.ProductName];
                            CleanDataA[nextCleanRow, (int)CleanCol.Quantity] = DirtyDataA[iRow, (int)DirtyCol.Quantity];
                            CleanDataA[nextCleanRow, (int)CleanCol.UnitPrice] = DirtyDataA[iRow, (int)DirtyCol.UnitPrice];
                            CleanDataA[nextCleanRow, (int)CleanCol.Vendor] = DirtyDataA[iRow, (int)DirtyCol.Vendor];
                            CleanDataA[nextCleanRow, (int)CleanCol.NetAmount] = DirtyDataA[iRow, (int)DirtyCol.NetAmount];
                            CleanDataA[nextCleanRow, (int)CleanCol.Currency] = DirtyDataA[iRow, (int)DirtyCol.Currency];
                            nextCleanRow++;
                            iRow++;

                            for (int i = 0; i < (int)RowTypeE.TotCol; i++)
                            {
                                RowDataA[i] = (DirtyDataA[iRow, i]);
                            }
                            rowType = RowType(RowDataA);
                        }
                        break;
                }
            }
            KAXL.NewSheetAndWriteCleanArr(CleanDataA,ws);

            KAXL.FormatAsTable("WarRoomData");

            WS newws = wb.ActiveSheet;

            int lastCol = (int)CleanCol.TotCol;

            newws.Cells[1,lastCol + 1].Value2 = "IncludeInMeeting";
            newws.Cells[1, lastCol + 2].Value2 = "Description";
            newws.Cells[1, lastCol + 3].Value2 = "Comments";
            newws.Cells[1, lastCol + 4].Value2 = "Status";
        }

        private static void UnmergeAll(WS ws, int LR, int LC)
        {
            for (int R = 1; R <= LR; R++)
            {
                for (int C = 1; C <= LC; C++)
                {
                    RG rgM = ws.Cells[R, C];
                    rgM.UnMerge();
                }
            }
        }

        public static int RowType(string[] RowDataA)
        {
            int rowType = 0;

            bool isReqID = IsReqID(RowDataA[2]);
            bool isTotal = IsTotal(RowDataA[2]);
            bool isBlank = IsBlank(RowDataA);
            if (isBlank)
                rowType = (int)RowTypeE.Blank;
            else if (!isBlank)
            {
                string v = RowDataA[1];
                if (v == "Purchase requisition statistics")
                    rowType = (int)RowTypeE.MainTitle;
                else if (v == "Company" || v == "Status")
                    rowType = (int)RowTypeE.MainTitleBlock;
                else if (v == "Employee")
                    rowType = (int)RowTypeE.Employee;
                else if (RowDataA[2] == "Purchase requisition ID")
                    rowType = (int)RowTypeE.ReqTitle;
                else if (isReqID)
                    rowType = (int)RowTypeE.ReqID;
                else if (RowDataA[2] == "Item number")
                    rowType = (int)RowTypeE.SubHeadings;
                else if (v == null && RowDataA[6] != null)
                    rowType = (int)RowTypeE.LineData;
                else if (v == null && RowDataA[2] == null && RowDataA[3] == null && RowDataA[4] == null &&
                        RowDataA[5] == null && RowDataA[15] != null)
                    rowType = (int)RowTypeE.RightTitleBlock;
                else if (RowDataA[17] == "CAD")
                    rowType = (int)RowTypeE.SecSubTotals;
                else if (isTotal)
                    rowType = (int)RowTypeE.GrandTotals;
            }
            return rowType;
        }
        private static bool IsReqID(string val)
        {
            bool reqID = false;

            if (val != null)
            {
                string sub;

                try
                {
                    sub = val.Substring(0, 2);
                }
                catch
                {
                    return false; 
                }

                if (sub == "PT")
                    reqID = true;
            }
            return reqID;
        }
        private static bool IsTotal(string val)
        {
            bool empName = false;

            if (val != null)
            {
                string sub;

                try
                {
                    sub = val.Substring(0, 3);
                }
                catch
                {
                    return false;
                }

                if (sub == "Tot")
                    empName = true;
            }
            return empName;
        }
        private static bool IsBlank(string[] data)
        {
            bool isBlank = true;

            for (int i = 0; i < 22; i++)
            {
                if (data[i] != null)
                {
                    isBlank = false;
                    break;
                }
                else
                    isBlank = true;
            }

            return isBlank;
        }
    }
}
