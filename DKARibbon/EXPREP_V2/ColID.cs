using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DKAExcelStuff;

namespace EXPREP_V2
{
    public class SourceColID
    {
        private ColIDL colIDL;
        // Classifies the column numbers for the "Open PO Lines" report that data is being read from
        public SourceColID(Worksheet ws) => colIDL = new ColIDL(ws);           

        public int VendorAccount => colIDL.GetColNum("Vendor account");
        public int PurchaseOrder => colIDL.GetColNum("Purchase order");
        public int LineNumber => colIDL.GetColNum("Line number");
        public int ItemNumber => colIDL.GetColNum("Item number");
        public int LineStatus => colIDL.GetColNum("Line status");
        public int Currency => colIDL.GetColNum("Currency");
        public int ProcurementCategory => colIDL.GetColNum("Procurement category");
        public int Site => colIDL.GetColNum("Site");
        public int Warehouse => colIDL.GetColNum("Warehouse");
        
        public int RevisedSchedDelDate => colIDL.GetColNum("Confirmed delivery date");
        public int CreatedDate => colIDL.GetColNum("Created date and time");
        public int OrigSchedDelDate => colIDL.GetColNum("Delivery date");

        public int Quantity => colIDL.GetColNum("Quantity");
        public int UnitPrice => colIDL.GetColNum("Unit price");
        public int NetAmount => colIDL.GetColNum("Net amount");
        public int AttentionInformation => colIDL.GetColNum("Attention information");
        public int ApprovalStatus => colIDL.GetColNum("Approval Status");
    }
    public class ExpRepColumn
    {
        private ColIDL colIDL;

        public ExpRepColumn(Worksheet ws) => colIDL = new ColIDL(ws);

        //POSource Class
        public int Requester => colIDL.GetColNum("Requester");
        public int POSourceType => colIDL.GetColNum("POSourceType");
        public int POSourceCode => colIDL.GetColNum("POSourceCode");
        public int Createdby => colIDL.GetColNum("CreatedBy");
        public int AttentionInfo => colIDL.GetColNum("AttentionInformation");
        public int Expeditor => colIDL.GetColNum("Expeditor");
        public int Direct => colIDL.GetColNum("Direct");

        // Item Class        
        public int ItemNumber => colIDL.GetColNum("ItemNumber");
        public int ItemDescription => colIDL.GetColNum("ItemDescription"); 

        // Vendor Class
        public int VendorAccount => colIDL.GetColNum("VendorAccount");
        public int VendorName => colIDL.GetColNum("VendorName");

        // Cash Class
        public int CAD => colIDL.GetColNum("TotalCAD");
        public int USD => colIDL.GetColNum("TotalUSD");
        public int UnitPriceCAD => colIDL.GetColNum("UnitPriceCAD");
        public int UnitPriceUSD => colIDL.GetColNum("UnitPriceUSD");
        public int Curr => colIDL.GetColNum("Currency");
        public int NetAmount => colIDL.GetColNum("Net Amount");

        // Dates Class
        public int Year => colIDL.GetColNum("Year");
        public int Month => colIDL.GetColNum("Month");
        public int Quarter => colIDL.GetColNum("Quarter");
        public int RecDate => colIDL.GetColNum("Actual Received Date"); 
        public int OriginalSchedDelDate => colIDL.GetColNum("OriginalScheduledDeliveryDate");
        public int RevisedSchedDelDate => colIDL.GetColNum("Current Delivery Date");
        public int POCreatedDate => colIDL.GetColNum("POCreatedDate");
        public int OnTime => colIDL.GetColNum("OnTime");
        public int DateAdded => colIDL.GetColNum("DateAdded");

        // No Class
        public int Entity => colIDL.GetColNum("Entity");
        public int ICO => colIDL.GetColNum("ICO");        
        public int ExpediteRequired => colIDL.GetColNum("ExpediteRequired");
        public int PONumber => colIDL.GetColNum("PO");
        public int LineNumber => colIDL.GetColNum("LineNumber");
        public int WH => colIDL.GetColNum("Warehouse");
        public int Quantity => colIDL.GetColNum("Quantity");        
        public int Status => colIDL.GetColNum("Status");
        public int Receiver => colIDL.GetColNum("Receiver");

        // Category Class
        public int Category => colIDL.GetColNum("Clean Category");

        public int totalColumnsInExpRep => colIDL.ColQ;
    }
}
