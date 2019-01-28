using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using Microsoft.Office.Tools.Ribbon;
using NR = Microsoft.Office.Tools.Excel.NamedRange;
using DKARibbon;

namespace DKAExcelStuff
{
    public class WarRoomPresent
    {
        private enum CleanDataColE
        {
            Nada,Employee, ReqID,ProjName,ProdName,Quantity,UnitPrice,VendorNum,NetAmount,Currency,Include,Description,Comments,Status,Num,Tot
        }
        private enum StatusE
        {
            Approved,RequiresFollowup,Tot
        }
        public static void PresentData()
        {
            var xlApp = new Application();

            WB wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            WS ws = wb.ActiveSheet;

            int LR = KAXL.LastRow(ws,1)+1;
            int LC = KAXL.LastCol(ws,1);

            List<string> RowDataL = new List<string>();

            bool includeLine = true;
            int compileCounter = 0; // used to bundle lines from the same reqID together, and insert into "Number" column

            string[,] dirtyArr = KAXL.LoadDirtyArr(ws, LR, (int)CleanDataColE.Tot);

            for (int r = 1; r < LR; r++)
            {
                includeLine = IncludeLine(dirtyArr[r, (int)CleanDataColE.Include]);

                if (includeLine)
                {
                    for (int c = 1; c < (int)CleanDataColE.Tot; c++)
                    {
                        if (c == (int)CleanDataColE.Num)
                        {
                            if (dirtyArr[r, (int)CleanDataColE.ReqID] == dirtyArr[(r - 1), (int)CleanDataColE.ReqID])
                            {
                                RowDataL.Add(compileCounter.ToString());
                            }
                            else
                            {
                                compileCounter++;
                                RowDataL.Add(compileCounter.ToString());
                            }
                        }
                        else
                        {
                            RowDataL.Add(dirtyArr[r, c]);
                        }
                            
                    }
                }
            }
            

            //numLines++; // inserting row at top for headings...
            int numCol = (int)CleanDataColE.Tot - 1;
            int numLines = RowDataL.Count() / numCol;

            string[,] CleanDataArr = new string[(numLines+1),numCol];
            int nextI = 0;

            for (int R = 0; R <= (numLines); R++)
            {
                if (R == 0)
                {
                    // print headings on first row of array
                    for (int CC = 1; CC < (int)CleanDataColE.Tot; CC++)
                    {
                        CleanDataArr[0, (CC-1)] = ((CleanDataColE)CC).ToString();
                    }
                }
                else
                {
                    for (int C = 0; C < numCol; C++)
                    {
                        CleanDataArr[R, C] = RowDataL[nextI];
                        nextI++;
                    }
                }
            }
            KAXL.NewSheetAndWriteCleanArr(CleanDataArr, ws);
            KAXL.FormatAsTable("WarRoomPresentation");

            WS wsPresenation = wb.ActiveSheet;
            
            RG rng = wsPresenation.Cells[(numLines + 3), 1];
            rng.Value2 = "WarRoom Meeting Date:";
            rng.Font.Bold = true;

            RG rng2 = wsPresenation.Cells[(numLines + 4), 1];
            rng.Value2 = "Attendees / Approvers:";

        }
        private static bool IncludeLine(string val)
        {
            bool includeLine = false;

            if (val == "Yes" ||
                val == "yes" ||
                val == "True" ||
                val == "true" ||
                val == "y" ||
                val == "Y")
                includeLine = true;

            return includeLine;
        }
    }
}
