using System.Collections.Generic;

namespace SmartOCR
{
    internal class TableTreeNodeContent : ITreeNodeContent
    // TODO: add summary.
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public List<long> Lines { get; set; }
        public string Name { get; set; }
        public string NodeLabel { get; set; }
        public string RE_Pattern { get; set; }
        public bool Status { get; set; }
        public string CheckValue { get; set; }
        public string FoundValue { get; set; }
        public string ValueType { get; set; }

        public TableTreeNodeContent()
        {
            Lines = new List<long>();
        }

        public TableTreeNodeContent(TableTreeNodeContent content) : this()
        {
            Row = content.Row;
            Name = content.Name;
            NodeLabel = content.NodeLabel;
            RE_Pattern = content.RE_Pattern;
            Status = content.Status;
            CheckValue = content.CheckValue;
            FoundValue = content.FoundValue;
            ValueType = content.ValueType;
            Column = content.Column;
        }
    }
}