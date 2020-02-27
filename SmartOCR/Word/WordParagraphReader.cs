namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Used to collect data from Word paragraphs.
    /// </summary>
    internal class WordParagraphReader
    {
        /// <summary>
        /// Defines lower bound of document text element length.
        /// </summary>
        private const byte MinimalTextLength = 2;

        private readonly Document document;
        private long pageIndex;
        private SortedDictionary<decimal, List<ParagraphContainer>> paragraphs;
        private List<int> paragraphIndexes;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordParagraphReader"/> class.
        /// </summary>
        /// <param name="document">A Word document.</param>
        /// <param name="pageIndex">Page index, from which paragraphs are read.</param>
        public WordParagraphReader(Document document, long pageIndex)
        {
            this.document = document;
            this.pageIndex = pageIndex;
            this.ResetParagraphIndexes();
        }

        /// <summary>
        /// Returns mapping between paragraph horizontal location and collection of paragraphs at location.
        /// </summary>
        /// <returns>A mapping between paragraph horizontal location and collection of paragraphs at location.</returns>
        public SortedDictionary<decimal, List<ParagraphContainer>> GetValidParagraphs()
        {
            return this.paragraphs ?? this.ReadDocument();
        }

        /// <summary>
        /// Returns mapping between paragraph horizontal location and collection of paragraphs at location.
        /// </summary>
        /// <param name="newPageIndex">New page index.</param>
        /// <returns>A mapping between paragraph horizontal location and collection of paragraphs at location.</returns>
        public SortedDictionary<decimal, List<ParagraphContainer>> GetValidParagraphs(long newPageIndex)
        {
            this.pageIndex = newPageIndex;
            this.ResetParagraphIndexes();
            return this.GetValidParagraphs();
        }

        private static void TryGetDataFromParagraph(List<ParagraphContainer> paragraphCollection, Range singleRange)
        {
            if (singleRange.Text.Length > MinimalTextLength)
            {
                paragraphCollection.Add(new ParagraphContainer(singleRange));
            }
        }

        private void AddParagraphToContents(List<ParagraphContainer> paragraphs, int paragraphIndex)
        {
            decimal location = paragraphs[paragraphIndex].VerticalLocation;
            if (!this.paragraphs.ContainsKey(location))
            {
                this.paragraphs.Add(location, new List<ParagraphContainer>());
            }

            this.paragraphs[location].Add(paragraphs[paragraphIndex]);
        }

        private SortedDictionary<decimal, List<ParagraphContainer>> ReadDocument()
        {
            this.paragraphs = new SortedDictionary<decimal, List<ParagraphContainer>>();
            List<ParagraphContainer> paragraphs = this.GetParagraphsOnPage();
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++)
            {
                this.AddParagraphToContents(paragraphs, paragraphIndex);
            }

            return this.paragraphs;
        }

        private void ResetParagraphIndexes()
        {
            int startPoint = 0;

            for (int index = 1; index <= this.document.Paragraphs.Count; index++)
            {
                long currentPage = this.document.Paragraphs[index].Range.Information[WdInformation.wdActiveEndPageNumber];
                if (currentPage > this.pageIndex)
                {
                    this.paragraphIndexes = Enumerable.Range(startPoint, index - startPoint + 1).ToList();
                    return;
                }

                if (startPoint == 0 && this.pageIndex == currentPage)
                {
                    startPoint = index;
                }
            }

            this.paragraphIndexes = Enumerable.Range(startPoint, this.document.Paragraphs.Count).ToList();
        }

        private List<ParagraphContainer> GetParagraphsOnPage()
        {
            var paragraphCollection = new List<ParagraphContainer>();
            for (int i = 0; i < this.paragraphIndexes.Count; i++)
            {
                TryGetDataFromParagraph(paragraphCollection, this.document.Paragraphs[this.paragraphIndexes[i]].Range);
            }

            return paragraphCollection;
        }
    }
}
