namespace SmartOCR.Word
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents mapping of multiple Word paragraphs to their horizontal location on document page.
    /// </summary>
    public class ParagraphMapping : SortedDictionary<decimal, List<ParagraphContainer>>
    {
    }
}
