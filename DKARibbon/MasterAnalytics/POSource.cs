using System;
using System.Collections.Generic;
using System.Text;

namespace DKAExcelStuff
{
    public class POSource
    {
        enum POSourceE { Unknown, ProdOrder, PReq, Project, AFE, MinMax, ICO }
        private readonly List<string> POSourcePrefixes = new List<string>()
            { "Nada", "PRO", "PRQ", "PJN", "AFE", "MIN", "ICO" };
        enum DataSplitE { POSource, Req, Buyer }

        private readonly POSource[] poSourceArray;

        public POSource() { }

        public POSource(string attInfo)
        {
            try
            {
                attInfo.Trim();

                // remove data after 3rd '/'
                string[] split = attInfo.Split('/');
                attInfo = null;
                attInfo += split[0] + "/" + split[1] + "/" + split[2];

                // get rid of spaces, & symbols (#, -, :) **********************************
                char[] charA = new char[attInfo.Length];
                charA = attInfo.ToCharArray();
                attInfo = null;

                IsMultiLinePO = IsMultiLine(charA);

                for (int i = 0; i < charA.Length; i++)
                {
                    if (!char.IsWhiteSpace(charA[i]) && charA[i].ToString() != "#" && charA[i].ToString() != ":" &&
                        charA[i].ToString() != "-")
                        attInfo += char.ToUpper(charA[i]);
                }

                // get rid of PTCA & PTUS ****************************************************
                charA = new char[attInfo.Length];
                charA = attInfo.ToCharArray();
                attInfo = null;
                int skipIndexes = 3;

                for (int i = 0; i < (charA.Length - skipIndexes); i++)
                {
                    if (!ScrubEntity(charA, i))
                        attInfo += charA[i];
                    else
                        i += skipIndexes; // to skip the entire 4 letter entity string
                }

                for (int i = 0; i < skipIndexes; i++)
                    attInfo += charA[charA.Length - (skipIndexes - i)];

                // split data
                string[] dataSplit = attInfo.Split('/');

                //determine how many POSources there are
                if (IsMultiLinePO)
                {
                    char[] charA2 = dataSplit[(int)DataSplitE.POSource].ToCharArray();
                    int QPOSources = 1;

                    for (int i = 0; i < charA2.Length; i++)
                    {
                        if (charA2[i].ToString() == ",")
                            QPOSources++;
                    }
                    poSourceArray = new POSource[QPOSources];
                    for (int i = 0; i < QPOSources; i++)
                    {
                        POSource pos = new POSource();
                        poSourceArray[i] = pos;
                    }
                }

                string prefix = dataSplit[(int)DataSplitE.POSource].Substring(0, 3);
                int enumIndex = GetPOSourceType(prefix);

                POSourceType = ((POSourceE)enumIndex).ToString();
                CreatedBy = dataSplit[(int)DataSplitE.Buyer];

                if (IsMultiLinePO)
                {
                    List<string> reqL = Requesters(dataSplit[(int)DataSplitE.Req]);
                    List<string> multiLinePOSourceCodesL = ScrubMultiLinePOSourceCodes(dataSplit[(int)DataSplitE.POSource]);

                    for (int i = 0; i < GetQPO; i++)
                    {
                        poSourceArray[i].POSourceType = POSourceType;
                        poSourceArray[i].CreatedBy = CreatedBy;
                        poSourceArray[i].POSourceCode = multiLinePOSourceCodesL[i];
                        poSourceArray[i].Requester = GetRequester(reqL, i);
                    }
                }
                else
                {
                    POSourceCode = ScrubPOSourceCode(enumIndex, dataSplit[(int)DataSplitE.POSource]);
                    Requester = dataSplit[(int)DataSplitE.Req];
                }
                AttInfo = attInfo;
            }
            catch
            {
                POSourceType = null;
                POSourceCode = null;
                Requester = null;
                CreatedBy = null;
                AttInfo = null;
            }
        }
        public string POSourceCode { get; set; }
        public string Requester { get; set; }
        public string CreatedBy { get; set; }
        public string POSourceType { get; set; }
        public string AttInfo { get; set; }
        public bool IsMultiLinePO { get; set; }

        // indexer stuff
        public POSource this[int i]
        {
            get => poSourceArray[i];
            set => poSourceArray[i] = value;
        }
        public int GetQPO => poSourceArray.Length;

        private bool ScrubEntity(char[] charA, int i)
        {
            string test = Convert.ToString(charA[i].ToString() + charA[i + 1].ToString()
                + charA[i + 2].ToString() + charA[i + 3].ToString());

            if (test == "PTCA" || test == "PTUS" || test == "HMCA" || test == "HMUS" || test == "PMEX")
                return true;
            else
                return false;
        }
        private string ScrubZeros(string code)
        {
            char[] charA = new char[code.Length];
            charA = code.ToCharArray();
            double zero = char.GetNumericValue('0');
            string combo = null;
            int firstPossibleZero = 3;
            int lastPossibleZero = 7;

            if (char.GetNumericValue(charA[firstPossibleZero]) == zero)
            {
                for (int j = 0; j < firstPossibleZero; j++)
                {
                    combo += charA[j];
                }
                for (int i = firstPossibleZero; i < lastPossibleZero; i++)
                {
                    if (char.GetNumericValue(charA[i]) != zero)
                    {
                        combo += charA[i];
                    }
                }
                for (int k = lastPossibleZero; k < code.Length; k++)
                {
                    combo += charA[k];
                }
                return combo;
            }
            return code;
        }
        private string ScrubPOSourceCode(int enumIndex, string source)
        {
            if (enumIndex == (int)POSourceE.Project)
            {
                char[] charAA = new char[source.Length];
                charAA = source.ToCharArray();
                string combo = null;
                int hyphenInsertIndex = source.Length - 3;
                for (int i = 0; i < hyphenInsertIndex; i++)
                {
                    combo += charAA[i];
                }
                combo += "-";
                for (int i = hyphenInsertIndex; i < source.Length; i++)
                {
                    combo += charAA[i];
                }
                return combo;
            }
            else if (enumIndex == (int)POSourceE.ProdOrder)
            {
                int num;
                if (Int32.TryParse(source.Substring(0, 1), out num))
                {
                    source = POSourcePrefixes[enumIndex] + source;
                }
                return ScrubZeros(source);
            }
            return source;
        }
        private List<string> ScrubMultiLinePOSourceCodes(string dataSplit)
        {
            string[] subDataSplit = dataSplit.Split(',');
            List<string> aiL = new List<string>(subDataSplit.Length);

            // see if there is a prefix, or just number...and then add prefix
            for (int i = 0; i < subDataSplit.Length; i++)
            {
                var val = subDataSplit[i];

                if (val is string)
                {
                    aiL.Add(ScrubPOSourceCode((int)POSourceE.ProdOrder, val));
                }
                else // if prefix is missing, just number is there
                {
                    aiL.Add(POSourcePrefixes[POSourcePrefixes.IndexOf(POSourceType)] + val.ToString());
                }
            }
            return aiL;
        }
        private bool IsMultiLine(char[] charA)
        {
            for (int i = 0; i < charA.Length; i++)
            {
                if (charA[i].ToString() == ",")
                    return true;
            }
            return false;
        }
        private int GetPOSourceType(string prefix) => POSourcePrefixes.IndexOf(prefix);
        private List<string> Requesters(string requesters)
        {
            char[] charA = new char[requesters.Length];
            charA = requesters.ToCharArray();

            int ReqQ = 1;

            for (int i = 0; i < charA.Length; i++)
            {
                if (charA[i].ToString() == ",")
                    ReqQ++;
            }
            List<string> reqL = new List<string>(ReqQ);
            string[] datasplit = requesters.Split(',');

            for (int i = 0; i < ReqQ; i++)
            {
                reqL.Add(datasplit[i]);
            }
            return reqL;
        }
        private string GetRequester(List<string> reqL, int i) => (reqL.Count == GetQPO) ? reqL[i] : reqL[0];
    }
    public class OldPOSource
    {
        //public enum POSourceItemsE { POSourceType, POSourceCode, Requester, CreatedBy }

        //private const string ProdPrefix = "PRO";
        //private const string ReqPrefix = "PRQ";
        //private const string ProjPrefix = "PJN";
        //private const string AFEPrefix = "AFE";
        //private const string MinMaxPrefix = "MIN";
        //private const int QReturnProps = 4;

        //private string POSourceTypePOS { get; set; }
        //private string POSourceCodePOS { get; set; }
        //private bool IsMultiLinePOS { get; set; }
        //private string RequesterPOS { get; set; }
        //private string CreatedByPOS { get; set; }
        //public int QPOSources { get; set; }

        //public string[,] POSourceA { get; set; }

        //public POSource(string attInfo)
        //{
        //    // scrubb empty spaces
        //    attInfo.Trim();

        //    // load into Character array
        //    char[] charA = attInfo.ToCharArray();
        //    int length = charA.Length;

        //    // set attInfo to null prior to scrubbing array
        //    attInfo = null;

        //    // if there is multiple sources of data
        //    QPOSources = 1;  // starts at 1, because 2 commas for 3 elements, 1 for 2...

        //    for (int i = 0; i < length; i++)
        //    {
        //        if (charA[i].ToString() == ",")
        //            IsMultiLinePOS = true;
        //        if (!char.IsWhiteSpace(charA[i]) && charA[i].ToString() != "#" && charA[i].ToString() != ":")
        //            attInfo += char.ToUpper(charA[i]);
        //    }

        //    string[] dataSplit = attInfo.Split('/');

        //    // determine how many POSources there are
        //    if (IsMultiLinePOS)
        //    {
        //        char[] charA2 = dataSplit[0].ToCharArray();
        //        length = charA2.Length;

        //        for (int i = 0; i < length; i++)
        //        {
        //            if (charA2[i].ToString() == ",")
        //                QPOSources++;
        //        }
        //    }
        //    string[,] poSourceA = new string[QReturnProps, QPOSources];
        //    string prefix = dataSplit[0].Substring(0, 3);
        //    poSourceA[(int)POSourceItemsE.POSourceType, 0] = prefix;

        //    if (!IsMultiLinePOS)
        //    {
        //        poSourceA[(int)POSourceItemsE.POSourceCode, 0] = dataSplit[0];
        //        poSourceA[(int)POSourceItemsE.Requester, 0] = dataSplit[1];
        //        poSourceA[(int)POSourceItemsE.CreatedBy, 0] = dataSplit[2];
        //    }
        //    else
        //    {
        //        string[] dataSplit2 = attInfo.Split('/');
        //        string[] sourcesSplit = dataSplit[0].Split(',');
        //        string[] requestersSplit = dataSplit[1].Split(',');
        //        string buyer = dataSplit[2];

        //        for (int i = 0; i < QPOSources; i++)
        //        {
        //            poSourceA[(int)POSourceItemsE.POSourceCode, i] = sourcesSplit[i].ToString();
        //            poSourceA[(int)POSourceItemsE.Requester, i] = requestersSplit[i].ToString();
        //            poSourceA[(int)POSourceItemsE.CreatedBy, i] = dataSplit[2];
        //            poSourceA[(int)POSourceItemsE.POSourceType, i] = POSourceTypePOS;
        //        }
        //    }
        //    POSourceA = poSourceA;
        //}
    }
}
