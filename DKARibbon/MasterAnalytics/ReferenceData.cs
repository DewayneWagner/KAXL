using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using RG = Microsoft.Office.Interop.Excel.Range;

namespace DKAExcelStuff
{
    //public class ReferenceData
    //{
    //    public List<Vendor> VendorL { get; set; }
    //    public List<ExRate> ExRateL { get; set; }
    //    public Dictionary<string, ItemDesc> ItemDict { get; set; }
    //    //public Dictionary<string,ItemDesc> ItemDictTest { get; set; }

    //    public ReferenceData(List<string> sourceItemNumsL)
    //    {      
    //        Application xl = new Application();
    //        Workbook wb = xl.Workbooks.Open(MasterAnalytics.MasterDataPath);
    //        Worksheet ws = wb.Sheets[1];
    //        Range xlRange = ws.UsedRange;

    //        // Load ExRates Data
    //        double year, month, exRate;
    //        string curr;
    //        int iRow = 2;
    //        bool done = false;

    //        ExRateL = new List<ExRate>();

    //        do
    //        {
               
    //            curr = Convert.ToString(ws.Cells[iRow, Col.Curr].Value2);
                
    //            if (string.IsNullOrWhiteSpace(curr))
    //            {
    //                done = true;
    //            }
    //            else
    //            {
    //                year = Convert.ToDouble(ws.Cells[iRow, Col.Year].Value2);
    //                month = Convert.ToDouble(ws.Cells[iRow, Col.Month].Value2);                    
    //                exRate = Convert.ToDouble(ws.Cells[iRow, Col.ExRate].Value2);
    //                ExRate exrate = new ExRate(year, month, curr, exRate);
    //                ExRateL.Add(exrate);
    //            }
    //            iRow++;

    //        } while (!done);
            
    //        //ExRateL = exRateL;

    //        /*******************************************************************************************
    //         * Load Vendor info
    //         * ***************************************************************************************/

    //        string vendNum = null, vendName;
    //        iRow = 2;
    //        VendorL = new List<Vendor>();

    //        done = false;

    //        do
    //        {
    //            vendNum = Convert.ToString(ws.Cells[iRow, Col.VendNum].Value2);

    //            if (string.IsNullOrWhiteSpace(vendNum))
    //            {
    //                done = true;
    //            }
    //            else
    //            {
    //                vendName = Convert.ToString(ws.Cells[iRow, Col.VendName].Value2);
    //                Vendor v = new Vendor(vendNum, vendName);
    //                VendorL.Add(v);
    //            }

    //            iRow++;

    //        } while(!done);

    //        /********************************************************************************************
    //         * ITEM DICTIONARY
    //         * *******************************************************************************************/

    //        done = false;
    //        iRow = 2;
    //        int LR = KAXL.LastRow(ws, Col.ItemNum);
    //        string itemDesc = null;
    //        string itemCat = null;
    //        string itemNum = null;
    //        int jRow = LR + 1;

    //        ItemDict = new Dictionary<string, ItemDesc>();

    //        do
    //        {
    //            itemNum = Convert.ToString(ws.Cells[iRow, Col.ItemNum].Value2);

    //            if (itemNum is string && string.IsNullOrWhiteSpace(itemNum) || sourceItemNumsL.Count == 0 || iRow > LR)
    //            {
    //                // this is where the leftover itemNum's would be written at the bottom of itemlist
    //                if (sourceItemNumsL.Count > 0)
    //                {
    //                    foreach (string item in sourceItemNumsL)
    //                    {
    //                        ws.Cells[jRow, Col.ItemNum].Value2 = item;
    //                        jRow++;
    //                    }
    //                }
    //                done = true;
    //            }
    //            else
    //            {
    //                if (sourceItemNumsL.Contains(itemNum))
    //                {
    //                    try
    //                    {
    //                        itemDesc = ws.Cells[iRow, Col.ItemDesc].Value2;
    //                        itemCat = ws.Cells[iRow, Col.ItemCat].Value2;
    //                    }
    //                    catch
    //                    {
    //                        var itemD = ws.Cells[iRow, Col.ItemDesc].Value2;

    //                        if (itemD is string)
    //                            itemDesc = itemD;
    //                        else
    //                            itemDesc = Convert.ToString(itemD);

    //                        var itemC = ws.Cells[iRow, Col.ItemCat].Value2;

    //                        if (itemC is string)
    //                            itemCat = itemC;
    //                        else
    //                            itemCat = Convert.ToString(itemC);
    //                    }
    //                    ItemDesc item = new ItemDesc(itemNum, itemDesc, itemCat);
    //                    ItemDict.Add(item.ItemN, item);
    //                    sourceItemNumsL.Remove(itemNum);
    //                }
    //            }
    //            iRow++;
    //        } while (!done);

    //        wb.Close(true);
    //        xl.Quit();
    //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
    //    }
    //}
    //public class ExRate
    //{
    //    public double ExRateYear { get; set; }
    //    public double ExRateMonth { get; set; }
    //    public string OriginCurrency { get; set; }
    //    public double ExRateToCAD { get; set; }

    //    public ExRate(double exRateYear, double exRateMonth, string originCurrency, double exRate)
    //    {
    //        ExRateYear = exRateYear;
    //        ExRateMonth = exRateMonth;
    //        OriginCurrency = originCurrency;
    //        ExRateToCAD = exRate;
    //    }
    //}
    //public class ItemDesc
    //{
    //    public string ItemN { get; set; }
    //    public string ItemD { get; set; }
    //    public string ItemCategory { get; set; }

    //    private List<ItemDesc> itemL = new List<ItemDesc>();

    //    public ItemDesc() { }

    //    public ItemDesc(string itemN, string itemD, string itemCategory)
    //    {
    //        ItemN = itemN;
    //        ItemD = itemD;
    //        ItemCategory = itemCategory;
    //    }
    //    public ItemDesc this[int i]
    //    {
    //        get => itemL[i];
    //        set => itemL[i] = value;
    //    }
    //}    
    //public class Vendor
    //{
    //    public string VendorCAccount { get; set; }
    //    public string VendorCName { get; set; }

    //    public Vendor(string vendorAccount, string vendorName)
    //    {
    //        VendorCAccount = vendorAccount;
    //        VendorCName = vendorName;
    //    }
    //}
    //public class Col
    //{
    //    public static int Curr = 1;
    //    public static int Year = 2;
    //    public static int Month = 3;
    //    public static int ExRate = 4;
    //    public static int ItemNum = 9;
    //    public static int ItemDesc = 10;
    //    public static int ItemCat = 11;
    //    public static int VendNum = 15;
    //    public static int VendName = 16;

    //    // TEST - REMOVE AFTER
    //    public static int testItemN = 20;
    //    public static int testItemD = 21;
    //    public static int testItemC = 22;
    //}
}
