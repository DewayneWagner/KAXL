using System;
using DKAExcelStuff;
using System.Collections.Generic;

namespace EXPREP_V2
{
    //public class Dates
    //{
    //    public Dates() { }

    //    public Dates(Master m)
    //    {
    //        ReceivedDate rd = new ReceivedDate(m);
    //        RevisedSchedDeliveryDate rsd = new RevisedSchedDeliveryDate(m);
    //    }

    //    public Dates(string revisedSchedDelDate, string origSchedDelDate, string poCreatedDate)
    //    {            
    //        OrigSchedDelDate = KAXL.ReadDateTime(origSchedDelDate);
    //        string dateTimeMinValueString = Convert.ToString(DateTime.MinValue);

    //        if (revisedSchedDelDate != null)
    //        {
    //            RevisedSchedDelDate = new RevisedSchedDeliveryDate(revisedSchedDelDate);
    //        }
    //        else if (revisedSchedDelDate == null)
    //        {
    //            RevisedSchedDelDate = new RevisedSchedDeliveryDate(dateTimeMinValueString);
    //        }

    //        if (OrigSchedDelDate == DateTime.MinValue)
    //        {
    //            OrigSchedDelDate = KAXL.ReadDateTime(revisedSchedDelDate);
    //        }
            
    //        POCreatedDate = KAXL.ReadDateTime(poCreatedDate);            
    //    }

    //    public RevisedSchedDeliveryDate RevisedSchedDelDate { get; set; }
    //    public DateTime OrigSchedDelDate { get; set; }
    //    public DateTime POCreatedDate { get; set; }
    //    public ReceivedDate ReceivedDate {get; set;}

    //    private TimeSpan onTimeCalc => ReceivedDate != null && OrigSchedDelDate != null ? 
    //        (ReceivedDate.Actual - OrigSchedDelDate) : TimeSpan.MinValue;

    //    public double OnTime => onTimeCalc != TimeSpan.MinValue ? Convert.ToDouble(onTimeCalc) : 0;

    //    public int Year => POCreatedDate.Year;
    //    public int Month => POCreatedDate.Month;
    //    public string Quarter => KAXL.Quarter(Month, Year);
    //}  
}
