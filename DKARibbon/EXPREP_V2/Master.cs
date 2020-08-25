using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DKAExcelStuff;
using WS = Microsoft.Office.Interop.Excel.Worksheet;

namespace EXPREP_V2
{
    public class Master
    {
        public enum SheetNamesE { Nada, ExpRep, Rules, Pivot, PTCA, PTUS, HMCA, MasterData }
        public enum MasterDataColumnsE {z,ExRateKey,ExRateCurrFrom,ExRateCurrTo,ExRateYear,ExRateMonth,ExRate,z1,
            z2,z3,z4,z5,ItemNum,ItemDesc,ItemCat,z6,z7,VendorAccount,VendorName }

        public Master() { }

        private frmEXPREP_V2_WINDOW form;

        public Master(KAXLApp kaxlapp, frmEXPREP_V2_WINDOW f)
        {
            form = f;
            stopWatch = new StopWatch();
            stopWatch.StartTime = DateTime.Now;

            updateMetrics = new UpdateMetrics();

            kaxlApp = kaxlapp;

            ExpRepColumn = new ExpRepColumn(kaxlApp.WB.Sheets[(int)SheetNamesE.ExpRep]);
            Dates = new AllDates(this);            
            VendorDict = new Vendor(this); // to initialize new vendordict
            ItemDict = new Item(this); // to initialize new itemdict
            ExRateDict = new ExRate(this);
            PODictionaryInExpRep = new PODictionaryInExpRep(this);            
            CategoryReferenceDictionary = new CategoryReferenceDictionary();
                        
            // start reading lines of data from the rawData (cycles between tabs)
            POLinesList = new ScrubbedPOLine(this);
            //AddToExpRep a = new AddToExpRep(this);

            WriteObjectArrayToExpRep = new WriteObjectArrayToExpRep(this);

            //WriteToExpRep = new WriteToExpRep(this, POLinesList.GetList());
        }

        public KAXLApp kaxlApp { get; set; }
        public ExRate ExRateDict { get; set; }
        public ExpRepColumn ExpRepColumn { get; set; }
        public Vendor VendorDict { get; set; }
        public Item ItemDict { get; set; }
        public ScrubbedPOLine POLinesList { get; set; }
        public PODictionaryInExpRep PODictionaryInExpRep { get; set; }        
        public CategoryReferenceDictionary CategoryReferenceDictionary {get; set;}
        public AllDates Dates { get; set; }
        public StopWatch stopWatch { get; set; }
        public UpdateMetrics updateMetrics { get; set; }

        public WriteObjectArrayToExpRep WriteObjectArrayToExpRep { get; }
        
        public class StopWatch
        {
            public StopWatch() { }

            public DateTime StartTime { get; set; }
            public DateTime EndTime { get; set; }
            public TimeSpan ElapsedTime => EndTime - StartTime;
        }
        public class UpdateMetrics
        {
            public UpdateMetrics() { }

            public int QTotalUpdatedLines { get; set; }
            public int QUpdatedReceivedDates { get; set; }
            public int QUpdatedRevisedDeliveryDates { get; set; }
            public int QItemDescriptionsUpdated { get; set; }
        }        
    }    
}
