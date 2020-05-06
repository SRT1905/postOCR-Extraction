namespace SmartOCR.Word.Read
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Word;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

    /// <summary>
    /// Used to collect data from Word paragraphCollection.
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
        private Range currentParagraphRange = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordParagraphReader"/> class.
        /// </summary>
        /// <param name="document">A Word document.</param>
        /// <param name="pageIndex">Page index, from which paragraphCollection are read.</param>
        public WordParagraphReader(Document document, int pageIndex)
        {
            this.document = document;
            this.pageIndex = pageIndex;
            this.currentParagraphRange = this.document.Paragraphs[1].Range;
        }

        /// <summary>
        /// Returns mapping between paragraph horizontal location and collection of paragraphCollection at location.
        /// </summary>
        /// <returns>A mapping between paragraph horizontal location and collection of paragraphCollection at location.</returns>
        public ParagraphMapping GetParagraphMapping()
        {
            Utilities.Debug($"Getting paragraphCollection from page {this.pageIndex}.", 3);
            return this.ReadDocument();
        }

        /// <summary>
        /// Returns mapping between paragraph horizontal location and collection of paragraphCollection at location.
        /// </summary>
        /// <param name="newPageIndex">New page index.</param>
        /// <returns>A mapping between paragraph horizontal location and collection of paragraphCollection at location.</returns>
        public ParagraphMapping GetParagraphMapping(int newPageIndex)
        {
            this.pageIndex = newPageIndex;
            while (this.currentParagraphRange != null && this.currentParagraphRange.Information[WdInformation.wdActiveEndPageNumber] < this.pageIndex)
            {
                this.currentParagraphRange = this.currentParagraphRange.Next(WdUnits.wdParagraph, 1);
            }

            return this.GetParagraphMapping();
        }

        /// <inheritdoc/>
        public List<WordTable> GetWordTables(int pageIndex)
        {
            this.pageIndex = pageIndex;
            return this.GetWordTables();
        }

        /// <inheritdoc/>
        public List<WordTable> GetWordTables()
        {
            var tables = new List<WordTable>();

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

        private void AddParagraphToContents(List<ParagraphContainer> paragraphCollection, int paragraphIndex)
        {
            decimal location = paragraphCollection[paragraphIndex].VerticalLocation;
            this.TryAddNewLocation(location);

            this.paragraphs[location].Add(paragraphCollection[paragraphIndex]);
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
            List<ParagraphContainer> pageParagraphs = this.GetParagraphsOnPage();
            for (int paragraphIndex = 0; paragraphIndex < pageParagraphs.Count; paragraphIndex++)
            {
                this.AddParagraphToContents(pageParagraphs, paragraphIndex);
            }

            return this.paragraphs;
        }

        private List<ParagraphContainer> GetParagraphsOnPage()
        {
            var paragraphCollection = new List<ParagraphContainer>();
            while (this.currentParagraphRange != null && this.currentParagraphRange.Information[WdInformation.wdActiveEndPageNumber] == this.pageIndex)
            {
                TryGetDataFromParagraph(paragraphCollection, this.currentParagraphRange);
                this.currentParagraphRange = this.currentParagraphRange.Next(WdUnits.wdParagraph, 1);
            }

            return paragraphCollection;
        }
    }
}
