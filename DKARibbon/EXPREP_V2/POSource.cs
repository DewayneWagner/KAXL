using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXPREP_V2
{
    //public class POSource
    //{
    //    public enum POSourceE { Unknown, ProdOrder, PReq, Project, AFE, MinMax, HM, ICO }

    //    private readonly List<string> POSourcePrefixes = new List<string>()
    //        { null, "PRO", "PRQ", "PJN", "AFE", "MIN", "ICO" };

    //    enum DataSplitE {POSource,Req,Buyer}
        
    //    public POSource() {}

    //    public POSource(string attInfo)
    //    {
    //        try
    //        {
    //            attInfo.Trim();

    //            // remove data after 3rd '/'
    //            string[] split = attInfo.Split('/');
    //            attInfo = null;
    //            attInfo += split[0] + "/" + split[1] + "/" + split[2];

    //            // get rid of spaces, & symbols (#, -, :) **********************************
    //            char[] charA = new char[attInfo.Length];
    //            charA = attInfo.ToCharArray();
    //            attInfo = null;

    //            IsMultiLinePO = IsMultiLine(charA);
                                
    //            for (int i = 0; i < charA.Length; i++)
    //            {
    //                if(!char.IsWhiteSpace(charA[i]) && charA[i].ToString() != "#" && charA[i].ToString() != ":" &&
    //                    charA[i].ToString() != "-")
    //                    attInfo += char.ToUpper(charA[i]);
    //            }

    //            // get rid of PTCA & PTUS ****************************************************
    //            charA = new char[attInfo.Length];
    //            charA = attInfo.ToCharArray();
    //            attInfo = null;
    //            int skipIndexes = 3;
                
    //            for (int i = 0; i < (charA.Length - skipIndexes); i++)
    //            {
    //                if (!ScrubEntity(charA, i))
    //                    attInfo += charA[i];
    //                else
    //                    i += skipIndexes; // to skip the entire 4 letter entity string
    //            }

    //            for (int i = 0; i < skipIndexes; i++)
    //                attInfo += charA[charA.Length - (skipIndexes - i)];
                
    //            // split data
    //            string[] dataSplit = attInfo.Split('/');

    //            //determine how many POSources there are
    //            if (IsMultiLinePO)
    //            {
    //                char[] charA2 = dataSplit[(int)DataSplitE.POSource].ToCharArray();
    //                int QPOSources = 1;

    //                for (int i = 0; i < charA2.Length; i++)
    //                {
    //                    if (charA2[i].ToString() == ",")
    //                        QPOSources++;
    //                }
    //                MultiLinePOArray = new POSource[QPOSources];
    //                for (int i = 0; i < QPOSources; i++)
    //                {
    //                    POSource ai = new POSource();
    //                    MultiLinePOArray[i] = ai;
    //                }
    //            }                

    //            string prefix = dataSplit[(int)DataSplitE.POSource].Substring(0, 3);
    //            int enumIndex = GetPOSourceType(prefix);
    //            if (enumIndex == -1)
    //                enumIndex = 0;

    //            POSourceType = (POSourceE)enumIndex;
    //            CreatedBy = dataSplit[(int)DataSplitE.Buyer];

    //            if (IsMultiLinePO)
    //            {
    //                List<string> reqL = Requesters(dataSplit[(int)DataSplitE.Req]);
    //                List<string> multiLinePOSourceCodesL = ScrubMultiLinePOSourceCodes(dataSplit[(int)DataSplitE.POSource]);

    //                for (int i = 0; i < GetQPO; i++)
    //                {
    //                    MultiLinePOArray[i].POSourceType = POSourceType;
    //                    MultiLinePOArray[i].CreatedBy = CreatedBy;
    //                    MultiLinePOArray[i].POSourceCode = multiLinePOSourceCodesL[i];
    //                    MultiLinePOArray[i].Requester = GetRequester(reqL, i);                           
    //                }
    //            }
    //            else
    //            {
    //                POSourceCode = ScrubPOSourceCode(enumIndex, dataSplit[(int)DataSplitE.POSource]);
    //                Requester = dataSplit[(int)DataSplitE.Req];
    //            }                
    //        }
    //        catch
    //        {
    //            POSourceType = POSourceE.Unknown;
    //            POSourceCode = null;
    //            Requester = null;
    //            CreatedBy = null;
    //            AttInfo = null;
    //        }            
    //    }
        
    //    public string POSourceCode { get; set; }
    //    public string Requester { get; set; }
    //    public string CreatedBy { get; set; }

    //    private POSourceE posourceType;
    //    public POSourceE POSourceType
    //    {
    //        get => posourceType;
    //        set => posourceType = value;
    //    } 
        
    //    public string AttInfo { get; set; }
    //    public bool IsMultiLinePO { get; set; }
    //    public POSource[] MultiLinePOArray { get; set; }

    //    public int GetQPO => MultiLinePOArray.Length;

    //    public POSource this[int i]
    //    {
    //        get => MultiLinePOArray[i];
    //        set => MultiLinePOArray[i] = value;
    //    }
    //    // this method tests if a 4 character sequence is an entity
    //    private bool ScrubEntity(char[] charA, int i)
    //    {
    //        string test = Convert.ToString(charA[i].ToString() + charA[i + 1].ToString() 
    //            + charA[i + 2].ToString() + charA[i + 3].ToString());

    //        if (test == "PTCA" || test == "PTUS" || test == "HMCA" || test == "HMUS" || test == "PMEX")
    //            return true;
    //        else
    //            return false;
    //    }
    //    // gets rid of leading 0's, not 0's in the project / pro number that need to stay
    //    private string ScrubZeros(string code)
    //    {
    //        char[] charA = new char[code.Length];
    //        charA = code.ToCharArray();
    //        double zero = char.GetNumericValue('0');
    //        string combo = null;
    //        int firstPossibleZero = 3;
    //        int lastPossibleZero = 7;

    //        if(char.GetNumericValue(charA[firstPossibleZero]) == zero)
    //        {
    //            for (int j = 0; j < firstPossibleZero; j++)
    //            {
    //                combo += charA[j];
    //            }
    //            for (int i = firstPossibleZero; i < lastPossibleZero; i++)
    //            {
    //                if (char.GetNumericValue(charA[i]) != zero)
    //                {
    //                    combo += charA[i];
    //                }                    
    //            }
    //            for (int k = lastPossibleZero; k < code.Length; k++)
    //            {
    //                combo += charA[k];
    //            }
    //            return combo;
    //        }
    //        return code;
    //    }

    //    private string ScrubPOSourceCode(int enumIndex, string source)
    //    {
    //        if (enumIndex == (int)POSourceE.Project)
    //        {
    //            char[] charAA = new char[source.Length];
    //            charAA = source.ToCharArray();
    //            string combo = null;
    //            int hyphenInsertIndex = source.Length - 3;
    //            for (int i = 0; i < hyphenInsertIndex; i++)
    //            {
    //                combo += charAA[i];
    //            }
    //            combo += "-";
    //            for (int i = hyphenInsertIndex; i < source.Length; i++)
    //            {
    //                combo += charAA[i];
    //            }
    //            return combo;
    //        }
    //        else if (enumIndex == (int)POSourceE.ProdOrder)
    //        {
    //            int num;
    //            if (Int32.TryParse(source.Substring(0,1),out num))
    //            {
    //                source = POSourcePrefixes[enumIndex] + source;
    //            }
    //            return ScrubZeros(source);
    //        }
    //        return source;
    //    }
    //    private List<string> ScrubMultiLinePOSourceCodes(string dataSplit)
    //    {
    //        string[] subDataSplit = dataSplit.Split(',');
    //        List<string> aiL = new List<string>(subDataSplit.Length);
            
    //        // see if there is a prefix, or just number...and then add prefix
    //        for (int i = 0; i < subDataSplit.Length; i++)
    //        {
    //            var val = subDataSplit[i];

    //            if(val is string)
    //            {
    //                aiL.Add(ScrubPOSourceCode((int)POSourceE.ProdOrder, val));
    //            }
    //            else // if prefix is missing, just number is there
    //            {                    
    //                aiL.Add(POSourcePrefixes[POSourcePrefixes.IndexOf(Convert.ToString(POSourceType))] + val.ToString());
    //            }
    //        }
    //        return aiL;
    //    }
    //    private bool IsMultiLine(char[] charA)
    //    {
    //        for (int i = 0; i < charA.Length; i++)
    //        {
    //            if (charA[i].ToString() == ",")
    //                return true;
    //        }
    //        return false;
    //    }
    //    private int GetPOSourceType(string prefix) => POSourcePrefixes.IndexOf(prefix);
    //    private List<string> Requesters(string requesters)
    //    {
    //        char[] charA = new char[requesters.Length];
    //        charA = requesters.ToCharArray();

    //        int ReqQ = 1;

    //        for (int i = 0; i < charA.Length; i++)
    //        {
    //            if(charA[i].ToString() == ",")
    //                ReqQ++;
    //        }
    //        List<string> reqL = new List<string>(ReqQ);
    //        string[] datasplit = requesters.Split(',');

    //        for (int i = 0; i < ReqQ; i++)
    //        {
    //            reqL.Add(datasplit[i]);
    //        }
    //        return reqL;
    //    }
    //    private string GetRequester(List<string> reqL,int i) => (reqL.Count == GetQPO)? reqL[i] : reqL[0];
    //}        
}
