﻿namespace SmartOCR
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Used to collect data from Word paragraphs.
    /// </summary>
    internal class WordParagraphReader : ITableReader
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
            List<WordTable> tables = new List<WordTable>(this.document.Tables.Count);

            for (int i = 1; i <= this.document.Tables.Count; i++)
            {
                Table table = this.document.Tables[i];
                if (table.Range.Information[WdInformation.wdActiveEndPageNumber] == this.pageIndex)
                {
                    tables.Add(new WordTable(table));
                }
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

        private void AddParagraphToContents(List<ParagraphContainer> paragraphs, int paragraphIndex)
        {
            decimal location = paragraphs[paragraphIndex].VerticalLocation;
            if (!this.paragraphs.ContainsKey(location))
            {
                this.paragraphs.Add(location, new List<ParagraphContainer>());
            }

            this.paragraphs[location].Add(paragraphs[paragraphIndex]);
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

            for (int index = 1; index <= this.document.Paragraphs.Count; index++)
            {
                int currentPage = this.document.Paragraphs[index].Range.Information[WdInformation.wdActiveEndPageNumber];
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
