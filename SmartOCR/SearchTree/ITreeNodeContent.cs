using System.Collections.Generic;

namespace SmartOCR
{
    internal interface ITreeNodeContent
    {
        string CheckValue { get; set; }
        string FoundValue { get; set; }
        List<long> Lines { get; set; }
        string Name { get; set; }
        string NodeLabel { get; set; }
        string RE_Pattern { get; set; }
        bool Status { get; set; }
        string ValueType { get; set; }
    }
}