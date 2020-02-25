using System;
using System.Collections.Generic;

namespace SmartOCR
{
    public class TableTreeNodeContent : ITreeNodeContent
    // TODO: add summary.
    {
        #region Properties
        public string CheckValue { get; set; }
        public int Column { get; set; }
        public string FoundValue { get; set; }
        public List<long> Lines { get;}
        public string Name { get; set; }
        public string NodeLabel { get; set; }
        public string RegExPattern { get; set; }
        public int Row { get; set; }
        public bool Status { get; set; }
        public string ValueType { get; set; }
        #endregion

        #region Constructors
        public TableTreeNodeContent()
        {
            Lines = new List<long>();
        }
        public TableTreeNodeContent(TableTreeNodeContent content) : this()
        {
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }
            Row = content.Row;
            Name = content.Name;
            NodeLabel = content.NodeLabel;
            RegExPattern = content.RegExPattern;
            Status = content.Status;
            CheckValue = content.CheckValue;
            FoundValue = content.FoundValue;
            ValueType = content.ValueType;
            Column = content.Column;
        }
        #endregion
    }
}