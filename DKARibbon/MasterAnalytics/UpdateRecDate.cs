using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XL = Microsoft.Office.Interop.Excel.Application;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace DKAExcelStuff
{
    //public class UpdateRecDate
    //{
    //    private List<string> HeadingsL = new List<string>() { };

    //    private enum PropsE {PONum,LineNum,Status,RecDate}

    //    public static void AddReceivedDate(Worksheet ws)
    //    {
    //        // to count number of received dates updated
    //        int updateCounter = 0;

    //        int LR = KAXL.LastRow(ws,1);
    //        List<Transaction> sourceL = new List<Transaction>();

    //        if (ws.Cells[1, 1].Value2 != "Purchase order" || ws.Cells[1, 2].Value2 != "Line number")
    //        {
    //            MessageBox.Show("Wrong columns, Jackass." + "\n" +
    //                "Column 1:  Purchase Order" + "\n" +
    //                "Column 2:  Line Number");
    //        }
    //        else
    //        {
    //            for (int i = 2; i <= LR; i++)
    //            {
    //                Transaction trans = new Transaction();
    //                sourceL.Add(trans);
    //                trans.PO = ws.Cells[i, 1].Value2;
    //                trans.LineNumber = Convert.ToInt32(ws.Cells[i, 2].Value2);
    //            }

    //            // open ExpReport
    //            XL xl = new XL();
    //            Workbook wb = xl.Workbooks.Open(@"R:\Supply Chain\ExpediteReport\ExpediteReport.xlsx");
    //            //Workbook wb = xl.Workbooks.Open(@"C:\Users\dewaynew\Desktop\Homework\ExpediteReport.xlsx");

    //            KAXL.CloseAndSaveWBIfOpen(wb);

    //            Worksheet expRep = wb.Sheets[1];

    //            if (expRep.AutoFilter != null)
    //                expRep.AutoFilterMode = false;

    //            int length = KAXL.LastCol(expRep,1);

    //            List<Transaction> expRepL = new List<Transaction>();

    //            DestColID dCol = new DestColID(expRep);

    //            DateTime today = DateTime.Today;
    //            DateTime recDate = today.AddDays(-1);

    //            var dayOfWeek = today.DayOfWeek;

    //            if (dayOfWeek == DayOfWeek.Monday)                
    //                recDate = today.AddDays(-3);                

    //            length = KAXL.LastRow(expRep,1);
    //            int listLength = sourceL.Count();

    //            for (int i = 2; i < length; i++)
    //            {
    //                if (expRep.Cells[i, dCol.Status].Value2 == "Open" && 
    //                    expRep.Cells[i, dCol.RecDate].Value2 == null &&
    //                    sourceL.Count != 0)
    //                {
    //                    Transaction trans = new Transaction();

    //                    trans.PO = expRep.Cells[i, dCol.PONumber].Value2;
    //                    trans.LineNumber = Convert.ToInt32(expRep.Cells[i,dCol.LineNumber].Value2);

    //                    for (int j = 2; j <= listLength; j++)
    //                    {
    //                        if (trans.PO == sourceL[(j-2)].PO && trans.LineNumber == sourceL[(j-2)].LineNumber)
    //                        {
    //                            expRep.Cells[i,dCol.Status].Value2 = "Closed";
    //                            expRep.Cells[i,dCol.ExpStatus].Value2 = "Physically Received";
    //                            updateCounter++;
                                
    //                            expRep.Cells[i,dCol.RecDate].Value2 = recDate.ToShortDateString();
    //                            sourceL.Remove(sourceL[(j-2)]);
    //                            listLength = sourceL.Count;

    //                            break;
    //                        }                            
    //                    }
    //                    if (listLength == 0)
    //                        break;
    //                }
    //            }
    //            KAXL.CloseApp(xl);
    //        }
    //        MessageBox.Show("Done, Jackass." + "\n" 
    //            + "Number of Received Dates Updated: " + updateCounter.ToString());
    //    }
    //}
}
