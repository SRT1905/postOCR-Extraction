namespace SmartOCR
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Used to collect data from Word paragraphs.
    /// </summary>
    public class WordParagraphReader : ITableReader
    {
        /// <summary>
        /// Defines lower bound of document text element length.
        /// </summary>
        private const byte MinimalTextLength = 2;

        private readonly Document document;
        private int pageIndex;
        private ParagraphMapping paragraphs;
        private List<int> paragraphIndexes;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordParagraphReader"/> class.
        /// </summary>
        /// <param name="document">A Word document.</param>
        /// <param name="pageIndex">Page index, from which paragraphs are read.</param>
        public WordParagraphReader(Document document, int pageIndex)
        {
            this.document = document;
            this.pageIndex = pageIndex;
            this.ResetParagraphIndexes();
        }

        /// <summary>
        /// Returns mapping between paragraph horizontal location and collection of paragraphs at location.
        /// </summary>
        /// <returns>A mapping between paragraph horizontal location and collection of paragraphs at location.</returns>
        public ParagraphMapping GetParagraphMapping()
        {
            Utilities.Debug($"Getting paragraphs from page {this.pageIndex}.", 3);
            return this.paragraphs ?? this.ReadDocument();
        }

        /// <summary>
        /// Returns mapping between paragraph horizontal location and collection of paragraphs at location.
        /// </summary>
        /// <param name="newPageIndex">New page index.</param>
        /// <returns>A mapping between paragraph horizontal location and collection of paragraphs at location.</returns>
        public ParagraphMapping GetParagraphMapping(int newPageIndex)
        {
            this.pageIndex = newPageIndex;
            this.ResetParagraphIndexes();
            return this.GetParagraphMapping();
        }

        /// <inheritdoc/>
        public List<WordTable> GetWordTables()
        {
            List<WordTable> tables = new List<WordTable>();

            for (int i = 1; i <= this.document.Tables.Count; i++)
            {
                this.TryAddTableToCollection(tables, this.document.Tables[i]);
            }

            return tables;
        }

        private static void TryGetDataFromParagraph(List<ParagraphContainer> paragraphCollection, Range singleRange)
        {
            if (singleRange.Text.Length > MinimalTextLength)
            {
                paragraphCollection.Add(new ParagraphContainer(singleRange));
            }
        }

        private static List<int> GetRange(int start, int finish)
        {
            return Enumerable.Range(start, finish).ToList();
        }

        private void AddParagraphToContents(List<ParagraphContainer> paragraphs, int paragraphIndex)
        {
            decimal location = paragraphs[paragraphIndex].VerticalLocation;
            this.TryAddNewLocation(location);

            this.paragraphs[location].Add(paragraphs[paragraphIndex]);
        }

        private void TryAddNewLocation(decimal location)
        {
            if (!this.paragraphs.ContainsKey(location))
            {
                this.paragraphs.Add(location, new List<ParagraphContainer>());
            }
        }

        private void TryAddTableToCollection(List<WordTable> tables, Table wordTable)
        {
            if (wordTable.Range.Information[WdInformation.wdActiveEndPageNumber] == this.pageIndex)
            {
                tables.Add(new WordTable(wordTable));
            }
        }

        private ParagraphMapping ReadDocument()
        {
            this.paragraphs = new ParagraphMapping();
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
            for (int paragraphIndex = 1; paragraphIndex <= this.document.Paragraphs.Count; paragraphIndex++)
            {
                int currentPage = this.document.Paragraphs[paragraphIndex].Range.Information[WdInformation.wdActiveEndPageNumber];
                if (currentPage > this.pageIndex)
                {
                    this.paragraphIndexes = GetRange(startPoint, paragraphIndex - startPoint + 1);
                    return;
                }

                startPoint = this.SetStartPoint(startPoint, paragraphIndex, currentPage);
            }

            this.paragraphIndexes = GetRange(startPoint, this.document.Paragraphs.Count);
        }

        private int SetStartPoint(int startPoint, int paragraphIndex, int currentPage)
        {
            return startPoint == 0 && this.pageIndex == currentPage
                ? paragraphIndex
                : startPoint;
        }

        private List<ParagraphContainer> GetParagraphsOnPage()
        {
            var paragraphCollection = new List<ParagraphContainer>();
            foreach (int paragraphIndex in this.paragraphIndexes)
            {
                TryGetDataFromParagraph(paragraphCollection, this.document.Paragraphs[paragraphIndex].Range);
            }

            return paragraphCollection;
        }
    }
}
