namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Performs collection and grouping of document contents.
    /// </summary>
    public sealed class WordReader : IDisposable // TODO: implement grid-like structure
    {
        /// <summary>
        /// Defines lower bound of document text element length.
        /// </summary>
        private const byte MinimalTextLength = 2;

        /// <summary>
        /// Defines limits of single document line.
        /// 72 points = 1 inch.
        /// </summary>
        private const byte VerticalPositionOffset = 6;

        /// <summary>
        /// Representation of Word document that is being read.
        /// </summary>
        private readonly Document document;
        private WordParagraphReader paragraphReader;
        private WordTextFrameReader frameReader;
        private List<WordTable> tables = new List<WordTable>();

        /// <summary>
        /// Initializes a new instance of the <see cref="WordReader"/> class.
        /// </summary>
        /// <param name="document">Word document representation.</param>
        public WordReader(Document document)
        {
            this.document = document;
        }

        /// <summary>
        /// Gets collection of processed Word tables.
        /// </summary>
        public List<WordTable> TableCollection
        {
            get
            {
                return this.tables;
            }

            private set
            {
                this.tables = value;
            }
        }

        /// <summary>
        /// Gets document contents grouped in separate lines.
        /// </summary>
        public LineMapping Mapping { get; private set; } = new LineMapping();

        /// <summary>
        /// Closes Word document if it is open.
        /// </summary>
        public void Dispose()
        {
            if (this.document != null)
            {
                WordApplication.CloseDocument(this.document);
            }
        }

        /// <summary>
        /// Performs data reading from all document pages.
        /// </summary>
        public void ReadDocument()
        {
            Utilities.Debug($"Reading document contents.", 1);
            int numberOfPages = this.document.Range().Information[WdInformation.wdNumberOfPagesInDocument];

            for (int pageIndex = 1; pageIndex <= numberOfPages; pageIndex++)
            {
                this.GetDataFromPage(pageIndex);
            }

            Utilities.Debug($"Total of {this.Mapping.Values.Sum(item => item.Count)} paragraphs were distributed into {this.Mapping.Count} lines.", 1);
        }

        /// <summary>
        /// Performs data reading from single document page.
        /// </summary>
        /// <param name="pageIndex">Index of page to read, starting from 1.</param>
        public void ReadDocument(int pageIndex)
        {
            if (pageIndex >= 1)
            {
                this.Mapping = this.ReadSinglePage(pageIndex);
            }
        }

        private static void AddDataFromCollectedParagraphs(ParagraphMapping documentContent, LineMapping newDocumentContent, int index)
        {
            AddDataFromSingleParagraph(
                documentContent,
                newDocumentContent,
                documentContent.Keys.ElementAt(index),
                newDocumentContent[newDocumentContent.Count][0].VerticalLocation);
        }

        private static ParagraphMapping MergeParagraphs(ParagraphMapping defaultParagraphMapping, ParagraphMapping frameParagraphMapping)
        {
            foreach (var item in frameParagraphMapping)
            {
                foreach (ParagraphContainer paragraph in item.Value)
                {
                    UpdateContentsWithParagraphs(defaultParagraphMapping, paragraph);
                }
            }

            return defaultParagraphMapping;
        }

        private static void AddDataFromSingleParagraph(
            ParagraphMapping documentContent,
            LineMapping newDocumentContent,
            decimal currentLocation,
            decimal previousLocation)
        {
            if (IsCurrentLocationWithinPreviousOne(currentLocation, previousLocation))
            {
                newDocumentContent[newDocumentContent.Count].AddRange(documentContent[currentLocation]);
            }
            else
            {
                newDocumentContent.Add(newDocumentContent.Count + 1, documentContent[currentLocation]);
            }
        }

        /// <summary>
        /// Groups document contents, arranged by vertical location on page, in separate lines.
        /// </summary>
        /// <param name="documentContent">Representation of read document contents.</param>
        /// <returns>An instance of <see cref="SortedDictionary{TKey, TValue}"/> where paragraphs are mapped to line.</returns>
        private static LineMapping GroupParagraphsByLine(ParagraphMapping documentContent)
        {
            LineMapping newDocumentContent = new LineMapping();

            if (documentContent.Count == 0)
            {
                return newDocumentContent;
            }

            newDocumentContent.Add(1, documentContent.Values.First());

            for (int i = 1; i < documentContent.Count; i++)
            {
                AddDataFromCollectedParagraphs(documentContent, newDocumentContent, i);
            }

            return SortParagraphs(newDocumentContent);
        }

        /// <summary>
        /// Adds ParagraphContainer instance to collection with maintaining sort order.
        /// </summary>
        /// <param name="paragraphCollection">Collection of ParagraphContainer objects.</param>
        /// <param name="textRangeContainer">ParagraphContainer instance to add.</param>
        /// <returns>Updated collection of ParagraphContainer objects.</returns>
        private static List<ParagraphContainer> InsertRangeInCollection(
            List<ParagraphContainer> paragraphCollection, ParagraphContainer textRangeContainer)
        {
            paragraphCollection.Add(textRangeContainer);
            paragraphCollection.Sort();
            return paragraphCollection;
        }

        private static bool IsCurrentLocationWithinPreviousOne(decimal currentLocation, decimal previousLocation)
        {
            return previousLocation - VerticalPositionOffset <= currentLocation &&
                   currentLocation <= previousLocation + VerticalPositionOffset;
        }

        private static LineMapping ShiftContentKeys(LineMapping pageContent, List<int> keys)
        {
            var shiftedMapping = new LineMapping();
            for (int i = 0; i < keys.Count; i++)
            {
                shiftedMapping.Add(keys[i] + 1, pageContent[keys[i]]);
            }

            pageContent = shiftedMapping;
            return pageContent;
        }

        private static LineMapping SortParagraphs(LineMapping newDocumentContent)
        {
            foreach (KeyValuePair<int, List<ParagraphContainer>> item in newDocumentContent)
            {
                item.Value.Sort();
            }

            return newDocumentContent;
        }

        private static void TryAddVerticalLocation(
            ParagraphMapping documentContent, ParagraphContainer container)
        {
            if (!documentContent.ContainsKey(container.VerticalLocation))
            {
                documentContent.Add(container.VerticalLocation, new List<ParagraphContainer>());
            }
        }

        private static void UpdateContentsWithParagraphs(
            ParagraphMapping documentContent, ParagraphContainer container)
        {
            if (container.Text.Length < MinimalTextLength)
            {
                return;
            }

            TryAddVerticalLocation(documentContent, container);
            documentContent[container.VerticalLocation] = InsertRangeInCollection(documentContent[container.VerticalLocation], container);
        }

        private void GetDataFromPage(int i)
        {
            var pageContent = this.ReadSinglePage(i);
            this.UpdateLineMapping(pageContent);
        }

        /// <summary>
        /// Gets grouped document content from specified page.
        /// </summary>
        /// <param name="pageIndex">Index of page to read.</param>
        /// <returns>Document contents, grouped by lines.</returns>
        private LineMapping ReadSinglePage(int pageIndex)
        {
            Utilities.Debug($"Getting contents from page {pageIndex}.", 2);
            this.InitializeReaders(pageIndex);
            ParagraphMapping documentContent = this.GetDataFromReaders(pageIndex);
            this.UpdateTableCollection();
            Utilities.Debug($"Distributing data from page {pageIndex}.", 2);
            return documentContent.Count == 0
                ? new LineMapping()
                : GroupParagraphsByLine(documentContent);
        }

        private void UpdateTableCollection()
        {
            foreach (ITableReader item in new ITableReader[2] { this.paragraphReader, this.frameReader })
            {
                this.TableCollection.AddRange(item.GetWordTables());
            }
        }

        private ParagraphMapping GetDataFromReaders(int pageIndex)
        {
            return MergeParagraphs(
                this.paragraphReader.GetParagraphMapping(pageIndex),
                this.frameReader.GetParagraphMapping());
        }

        private void InitializeReaders(int pageIndex)
        {
            this.paragraphReader = this.paragraphReader ?? new WordParagraphReader(this.document, pageIndex);
            this.frameReader = this.frameReader ?? new WordTextFrameReader(this.document, pageIndex);
        }

        /// <summary>
        /// Merges LineMapping dictionary with provided page content.
        /// </summary>
        /// <param name="pageContent">Grouped content from single page.</param>
        private void UpdateLineMapping(LineMapping pageContent)
        {
            if (pageContent.Count == 0)
            {
                return;
            }

            List<int> keys = pageContent.Keys.ToList();
            if (pageContent.ContainsKey(0))
            {
                pageContent = ShiftContentKeys(pageContent, keys);
            }

            if (this.Mapping.Count == 0)
            {
                this.Mapping = pageContent;
                return;
            }

            Utilities.Debug($"Merging collected data from current page with data from previous pages.", 2);
            this.UpdateLineMappingByEndLine(pageContent, keys);
        }

        private void UpdateLineMappingByEndLine(LineMapping pageContent, List<int> keys)
        {
            int endLine = this.Mapping.Keys.Last();
            for (int i = 0; i < keys.Count; i++)
            {
                int key = keys[i];
                this.Mapping.Add(key + endLine, pageContent[key]);
            }
        }
    }
}