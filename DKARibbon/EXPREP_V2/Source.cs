using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXPREP_V2
{
    public class Source
    {
        public enum SourceType { Unknown, ProdOrder, PReq, Project, AFE, MinMax, HM, ICO }
        private enum SourceTypePrefixes {Null, PRO, PRQ, PJN, AFE, MIN, ICO,Total }
        private enum DataSplitSection { Source,Requester,Creator,Scrap }

        public Source() { }

        private List<Source> _multiLineSourceList;

        private string[] _datasplit;
        //private List<string> _datasplit;
        private string[] _requesterDataSplit;

        public Source(string attentionInfo)
        {
            if(attentionInfo == null)
            {
                OriginalAttentionInfo = null;
                IsMultiLinePO = false;
                CreatedBy = null;
                Requester = null;
                Type = SourceType.Unknown;
                Code = null;
            }
            else
            {
                OriginalAttentionInfo = attentionInfo;
                _datasplit = attentionInfo.Split('/');

                if (DetermineIfMultiLinePO())
                {
                    IsMultiLinePO = true;
                    _multiLineSourceList = new List<Source>();

                    string[] _sourceDataSplit = _datasplit[(int)DataSplitSection.Source].Split(',');

                    bool isMultipleRequesters = IsMultipleRequesters();

                    if (isMultipleRequesters)
                    {
                        _requesterDataSplit = _datasplit[(int)DataSplitSection.Requester].Split(',');
                    }

                    for (int i = 0; i < _sourceDataSplit.Length; i++)
                    {
                        try
                        {
                            _multiLineSourceList.Add(new Source()
                            {
                                Requester = isMultipleRequesters ? _requesterDataSplit[i] : _datasplit[(int)DataSplitSection.Requester].ToUpper(),
                                CreatedBy = _datasplit[(int)DataSplitSection.Creator].ToUpper(),
                                Type = GetSourceType(),
                                Code = Type == SourceType.ProdOrder ? ScrubCode(_sourceDataSplit[i]) : _sourceDataSplit[i],
                            });
                        }
                        catch
                        {
                            OriginalAttentionInfo = null;
                            IsMultiLinePO = false;
                            CreatedBy = null;
                            Requester = null;
                            Type = SourceType.Unknown;
                            Code = null;
                        }
                    }
                }
                else
                {
                    try
                    {
                        IsMultiLinePO = false;
                        CreatedBy = _datasplit[(int)DataSplitSection.Creator].ToUpper();
                        Requester = _datasplit[(int)DataSplitSection.Requester].ToUpper();
                        Type = GetSourceType();
                        Code = Type == SourceType.ProdOrder ? ScrubCode(_datasplit[(int)DataSplitSection.Source]) : _datasplit[(int)DataSplitSection.Source];
                    }
                    catch
                    {
                        OriginalAttentionInfo = null;
                        IsMultiLinePO = false;
                        CreatedBy = null;
                        Requester = null;
                        Type = SourceType.Unknown;
                        Code = null;
                    }
                    
                }
            }            
        }

        public bool IsMultiLinePO { get; set; }
        public string OriginalAttentionInfo { get; }
        public SourceType Type { get; set; }
        public string Code { get; set; }
        public string CreatedBy { get; set; }
        public string Requester { get; set; }
        
        public Source this[int i]
        {
            get => _multiLineSourceList[i];
            set => _multiLineSourceList[i] = value;
        }
        public int QSourcesInList => _multiLineSourceList.Count;

        // test if multiLinePO
        private bool DetermineIfMultiLinePO()
        {
            char[] c = _datasplit[(int)DataSplitSection.Source].ToCharArray();

            bool isMultiLinePO = false;
            
            for (int i = 0; i < c.Length; i++)
            {
                isMultiLinePO = c[i].ToString() == "," ? true : false;
            }
            
            return isMultiLinePO;
        }
        // only going to scrub Production Order codes - leave entire string for other types of codes
        private string ScrubCode(string code)
        {
            char[] c = code.ToCharArray();
            code = null;

            for (int i = 4; i >= 1; i--)
            {
                code += c[c.Length - i];
            }

            return "PRO" + code;
        }
        private SourceType GetSourceType()
        {
            int q = 3;
            char[] c = _datasplit[(int)DataSplitSection.Source].ToCharArray();
            string testString;

            for (int i = 0; i < (c.Length - q); i++)
            {
                testString = (c[i] + c[i + 1] + c[i + 2]).ToString();
                for (int j = 0; j < (int)SourceTypePrefixes.Total; j++)
                {
                    if(testString == Convert.ToString((SourceTypePrefixes)j))
                    {
                        return (SourceType)j;
                    }
                }
            }
            return SourceType.Unknown;
        }
        private bool IsMultipleRequesters()
        {
            char[] c = _datasplit[(int)DataSplitSection.Requester].ToCharArray();

            foreach (char i in c)
            {
                if (i == ',')
                    return true;
            }
            return false;
        }
    }
}
