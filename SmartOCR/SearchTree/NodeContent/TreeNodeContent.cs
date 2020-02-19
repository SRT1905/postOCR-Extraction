using System;
using System.Collections.Generic;

namespace SmartOCR
{
    public class TreeNodeContent : ITreeNodeContent
    // TODO: add summary.
    {
        public decimal HorizontalParagraph { get; set; }
        public List<long> Lines { get;}
        public string Name { get; set; }
        public string NodeLabel { get; set; }
        public string RegExPattern { get; set; }
        public bool Status { get; set; }
        public string CheckValue { get; set; }
        public string FoundValue { get; set; }
        public string ValueType { get; set; }
        public int HorizontalStatus { get; set; }
        public int LineOffset { get; set; }

        public TreeNodeContent()
        {
            Lines = new List<long>();
        }

        public TreeNodeContent(TreeNodeContent content) : this()
        {
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }
            HorizontalParagraph = content.HorizontalParagraph;
            Name = content.Name;
            NodeLabel = content.NodeLabel;
            RegExPattern = content.RegExPattern;
            Status = content.Status;
            CheckValue = content.CheckValue;
            FoundValue = content.FoundValue;
            ValueType = content.ValueType;
            HorizontalStatus = content.HorizontalStatus;
            LineOffset = content.LineOffset;
        }
    }
}