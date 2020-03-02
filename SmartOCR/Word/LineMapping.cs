namespace SmartOCR
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents mapping of multiple Word paragraphs to their line location on document page.
    /// </summary>
    public class LineMapping : SortedDictionary<int, List<ParagraphContainer>>
    {
    }
}
