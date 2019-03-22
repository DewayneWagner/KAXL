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
            errorTracker = new ErrorTracker();

            kaxlApp = kaxlapp;

            // get the column numbers from expedite report
            ExpRepColumn = new ExpRepColumn(kaxlApp.WB.Sheets[(int)SheetNamesE.ExpRep]);

            // initialize classes to give access to master
            Cash c = new Cash(this);
            POLine p = new POLine(this);
            Category cat = new Category(this);

            // Properties to initialize
            Dates = new Dates(this);

            // load dictionaries
            ReceivedDateList = new ReceivedDateList(this);
            RevisedSchedDelDatesToUpdate = new RevisedSchedDelDatesToUpdate(this);

            VendorDict = new Vendor(this); // to initialize new vendordict
            ItemDict = new Item(this); // to initialize new itemdict
            ExRateDict = new ExRate(this);

            PODictionaryInExpRep = new PODictionaryInExpRep(this);
            
            CategoryReferenceDictionary = new CategoryReferenceDictionary(this);
                        
            // start reading lines of data from the rawData (cycles between tabs)
            POLinesList = new POLinesList(this);
            AddToExpRep a = new AddToExpRep(this);
        }

        public Category Category { get; set; }
        public KAXLApp kaxlApp { get; set; }
        public ExRate ExRateDict { get; set; }
        public SourceColID SColID { get; set; }
        public ExpRepColumn ExpRepColumn { get; set; }
        public Vendor VendorDict { get; set; }
        public Item ItemDict { get; set; }
        public POLinesList POLinesList { get; set; }
        public PODictionaryInExpRep PODictionaryInExpRep { get; set; }        
        public CategoryReferenceDictionary CategoryReferenceDictionary {get; set;}
        public ReceivedDateList ReceivedDateList { get; set; }
        public RevisedSchedDelDatesToUpdate RevisedSchedDelDatesToUpdate { get; set; }
        public StopWatch stopWatch { get; set; }
        public UpdateMetrics updateMetrics { get; set; }
        public ErrorTracker errorTracker { get; set; }

        // to give access to revised & received dates to update
        public Dates Dates { get; set; }
        
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
        }
        public class ErrorTracker
        {
            public string Process { get; set; }
            public string LineNumber { get; set; }

            public string GetErrorMessage() => "Error Occurred during " + Process + " process, " + "\n" +
                "on line number " + LineNumber;
        }
    }    
}
