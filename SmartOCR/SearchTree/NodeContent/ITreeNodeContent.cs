using System.Collections.Generic;

namespace SmartOCR
{
    public interface ITreeNodeContent   
    {
        #region Properties
        string CheckValue { get; set; }
        string FoundValue { get; set; }
        List<long> Lines { get; }
        string Name { get; set; }
        string NodeLabel { get; set; }
        string RegExPattern { get; set; }
        bool Status { get; set; }
        string ValueType { get; set; }
        #endregion
    }
}