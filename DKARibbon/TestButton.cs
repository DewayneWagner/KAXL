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

namespace DKAExcelStuff
{
    public static class TestButton
    {
        public static void TestM(KAXLApp kaxlApp)
        {
            DateTimeTest(kaxlApp);
        }
        public static void DateTimeTest(KAXLApp kaxl)
        {
            int sRow = 1, lRow = 5, iCol = 1;
            DateTime revDate;

            for (int i = sRow; i <= lRow; i++)
            {
                var rd = kaxl.WS.Cells[i,iCol].Value;
                revDate = (rd is DateTime) ? rd : Convert.ToDateTime(rd);
            }
            KAXLApp.CloseSheet(kaxl);
        }


        //public static void AttentionInfo(Worksheet ws)
        //{
        //    int firstRow = 2;
        //    int lastRow = 97;
        //    int cScrubbedAttInfo = 2;
        //    int cPOSourceCode = 3;
        //    int cRequester = 4;
        //    int cCreatedBy = 5;
        //    int cSourceType = 6;

        //    for (int i = firstRow; i <= lastRow; i++)
        //    {
        //        string attInfo = ws.Cells[i, 1].Value2;
        //        AttentionInfo attentionInfo = new AttentionInfo(attInfo);

        //        if (attentionInfo.IsMultiLinePO)
        //        {
        //            for (int ii = 0; ii < attentionInfo.GetQPO; ii++)
        //            {
        //                AttentionInfo ai = attentionInfo[ii];

        //                if(ii == 0)
        //                    ws.Cells[i, (cScrubbedAttInfo + ii)].Value2 = ai.AttInfo;

        //                ws.Cells[i, (cPOSourceCode + (4 * ii))].Value2 = ai.POSourceCode;
        //                ws.Cells[i, (cRequester + (4 * ii))].Value2 = ai.Requester;
        //                ws.Cells[i, (cCreatedBy+ (4 * ii))].Value2 = ai.CreatedBy;
        //                ws.Cells[i, (cSourceType+ (4 * ii))].Value2 = ai.POSourceType;
        //            }
        //        }
        //        else
        //        {
        //            ws.Cells[i, cScrubbedAttInfo].Value2 = attentionInfo.AttInfo;
        //            ws.Cells[i, cPOSourceCode].Value2 = attentionInfo.POSourceCode;
        //            ws.Cells[i, cRequester].Value2 = attentionInfo.Requester;
        //            ws.Cells[i, cCreatedBy].Value2 = attentionInfo.CreatedBy;
        //            ws.Cells[i, cSourceType].Value2 = attentionInfo.POSourceType;
        //        }
        //    }
        //}
        //public static void ReadItemsXL(Worksheet ws)
        //{
        //    int iRow = 2;
        //    int LR = KAXL.LastRow(ws, Col.ItemNum);
        //    string itemDesc = null;
        //    string itemCat = null;
        //    string itemNum = null;
        //    int jRow = LR + 1;

        //    string path = @"TestBinaryWrite.bin";

        //    Dictionary<string,ItemDesc> ItemDictTest = new Dictionary<string, ItemDesc>();
        //    Dictionary<string, ItemDesc> ItemDictBin = new Dictionary<string, ItemDesc>();

        //    for (int i = 2; i <= LR; i++)
        //    {
        //        itemNum = Convert.ToString(ws.Cells[i, Col.ItemNum].Value2);
        //        itemDesc = Convert.ToString(ws.Cells[i, Col.ItemDesc].Value2);
        //        itemCat = Convert.ToString(ws.Cells[i, Col.ItemCat].Value2);

        //        ItemDesc item = new ItemDesc(itemNum, itemDesc, itemCat);

        //        ItemDictTest.Add(itemNum, item);
        //    }

        //    foreach (KeyValuePair<string,ItemDesc> item in ItemDictTest)
        //    {
        //        ws.Cells[iRow, Col.testItemN].Value2 = item.Value.ItemN;
        //        ws.Cells[iRow, Col.testItemD].Value2 = item.Value.ItemD;
        //        ws.Cells[iRow, Col.testItemC].Value2 = item.Value.ItemCategory;
        //        iRow++;
        //    }

        //    BinaryWriter bw = new BinaryWriter(
        //        new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write));

        //    foreach (KeyValuePair<string,ItemDesc> item in ItemDictTest)
        //    {
        //        bw.Write(item.Value.ItemN);
        //        bw.Write(item.Value.ItemD);
        //        bw.Write(item.Value.ItemCategory);
        //    }

        //    bw.Close();

        //    BinaryReader br = new BinaryReader(
        //        new FileStream(path, FileMode.OpenOrCreate, FileAccess.Read));

        //    while (br.PeekChar() != (-1))
        //    {
        //        itemNum = br.ReadString();
        //        itemDesc = br.ReadString();
        //        itemCat = br.ReadString();

        //        ItemDesc item = new ItemDesc(itemNum, itemDesc, itemCat);
        //        ItemDictBin.Add(itemNum, item);
        //    }
        //    br.Close();
        //}
    }  
}