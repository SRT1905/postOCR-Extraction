namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Performs collection and grouping of document contents.
    /// </summary>
    public sealed class WordReader : IDisposable
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

        /// <summary>
        /// Initializes a new instance of the <see cref="WordReader"/> class.
        /// </summary>
        /// <param name="document">Word document representation.</param>
        public WordReader(Document document)
        {
            this.document = document;
            this.TableCollection = this.GetTables();
        }

        /// <summary>
        /// Gets collection of processed Word tables.
        /// </summary>
        public List<WordTable> TableCollection { get; private set; }

        /// <summary>
        /// Gets document contents grouped in separate lines.
        /// </summary>
        public SortedDictionary<long, List<ParagraphContainer>> LineMapping { get; private set; } = new SortedDictionary<long, List<ParagraphContainer>>();

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
            long numberOfPages = this.document.Range().Information[WdInformation.wdNumberOfPagesInDocument];

            for (int i = 1; i <= numberOfPages; i++)
            {
                this.GetDataFromPage(i);
            }
        }

        /// <summary>
        /// Performs data reading from single document page.
        /// </summary>
        /// <param name="pageIndex">Index of page to read, starting from 1.</param>
        public void ReadDocument(long pageIndex)
        {
            if (pageIndex >= 1)
            {
                this.LineMapping = this.ReadSinglePage(pageIndex);
            }
        }

        private static void AddDataFromCollectedParagraphs(SortedDictionary<decimal, List<ParagraphContainer>> documentContent, SortedDictionary<long, List<ParagraphContainer>> newDocumentContent, int index)
        {
            decimal currentLocation = documentContent.Keys.ElementAt(index);
            decimal previousLocation = newDocumentContent[newDocumentContent.Count][0].VerticalLocation;

            AddDataFromSingleParagraph(documentContent, newDocumentContent, currentLocation, previousLocation);
        }

        /// <summary>
        /// Gets data from TextFrame objects and adds it to document contents container.
        /// </summary>
        /// <param name="documentContent">Representation of read document contents.</param>
        /// <param name="frameCollection">Collection of TextFrame objects.</param>
        /// <returns>Representation of document contents that is extended by TextFrame objects.</returns>
        private static SortedDictionary<decimal, List<ParagraphContainer>> AddDataFromFrames(SortedDictionary<decimal, List<ParagraphContainer>> documentContent, List<TextFrame> frameCollection)
        {
            for (int i = 0; i < frameCollection.Count; i++)
            {
                documentContent = AddDataFromSingleFrame(documentContent, frameCollection[i]);
            }

            return documentContent;
        }

        /// <summary>
        /// Gets data from TextFrame object and adds it to document contents representation.
        /// </summary>
        /// <param name="documentContent">Representation of read document contents.</param>
        /// <param name="textFrame">TextFrame object containing text.</param>
        /// <returns>Representation of read document contents, extended by TextFrame contents.</returns>
        private static SortedDictionary<decimal, List<ParagraphContainer>> AddDataFromSingleFrame(SortedDictionary<decimal, List<ParagraphContainer>> documentContent, TextFrame textFrame)
        {
            List<ParagraphContainer> paragraphContainers = GetParagraphsFromTextFrame(textFrame);
            for (int i = 0; i < paragraphContainers.Count; i++)
            {
                UpdateContentsWithParagraphs(documentContent, paragraphContainers[i]);
            }

            return documentContent;
        }

        private static void AddDataFromSingleParagraph(SortedDictionary<decimal, List<ParagraphContainer>> documentContent, SortedDictionary<long, List<ParagraphContainer>> newDocumentContent, decimal currentLocation, decimal previousLocation)
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
        /// Extracts paragraphs from <see cref="TextFrame"/> object.
        /// </summary>
        /// <param name="textFrame">A <see cref="TextFrame"/> instance that contains text.</param>
        /// <returns>A collection of <see cref="ParagraphContainer"/> objects.</returns>
        private static List<ParagraphContainer> GetParagraphsFromTextFrame(TextFrame textFrame)
        {
            var paragraphContainers = new List<ParagraphContainer>();
            for (int i = 1; i <= textFrame.TextRange.Paragraphs.Count; i++)
            {
                paragraphContainers.Add(new ParagraphContainer(textFrame.TextRange.Paragraphs[i].Range));
            }

            return paragraphContainers;
        }

        /// <summary>
        /// Groups document contents, arranged by vertical location on page, in separate lines.
        /// </summary>
        /// <param name="documentContent">Representation of read document contents.</param>
        /// <returns>An instance of <see cref="SortedDictionary{TKey, TValue}"/> where paragraphs are mapped to line.</returns>
        private static SortedDictionary<long, List<ParagraphContainer>> GroupParagraphsByLine(SortedDictionary<decimal, List<ParagraphContainer>> documentContent)
        {
            if (documentContent.Count == 0)
            {
                return new SortedDictionary<long, List<ParagraphContainer>>();
            }

            var newDocumentContent = new SortedDictionary<long, List<ParagraphContainer>>()
            {
                { 1, documentContent.Values.First() },
            };

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
        private static List<ParagraphContainer> InsertRangeInCollection(List<ParagraphContainer> paragraphCollection, ParagraphContainer textRangeContainer)
        {
            paragraphCollection.Add(textRangeContainer);
            paragraphCollection.Sort();
            return paragraphCollection;
        }

        private static bool IsCurrentLocationWithinPreviousOne(decimal currentLocation, decimal previousLocation)
        {
            return previousLocation - VerticalPositionOffset <= currentLocation && currentLocation <= previousLocation + VerticalPositionOffset;
        }

        private static SortedDictionary<long, List<ParagraphContainer>> ShiftContentKeys(SortedDictionary<long, List<ParagraphContainer>> pageContent, List<long> keys)
        {
            var shiftedMapping = new SortedDictionary<long, List<ParagraphContainer>>();
            for (int i = 0; i < keys.Count; i++)
            {
                shiftedMapping.Add(keys[i] + 1, pageContent[keys[i]]);
            }

            pageContent = shiftedMapping;
            return pageContent;
        }

        private static SortedDictionary<long, List<ParagraphContainer>> SortParagraphs(SortedDictionary<long, List<ParagraphContainer>> newDocumentContent)
        {
            foreach (KeyValuePair<long, List<ParagraphContainer>> item in newDocumentContent)
            {
                item.Value.Sort();
            }

            return newDocumentContent;
        }

        private static void TryAddFrame(List<TextFrame> frames, Shape shape)
        {
            if (shape.TextFrame != null && shape.TextFrame.HasText != 0 && shape.TextFrame.TextRange.Text.Length > MinimalTextLength)
            {
                frames.Add(shape.TextFrame);
            }
        }

        private static void TryAddVerticalLocation(SortedDictionary<decimal, List<ParagraphContainer>> documentContent, ParagraphContainer container)
        {
            if (!documentContent.ContainsKey(container.VerticalLocation))
            {
                documentContent.Add(container.VerticalLocation, new List<ParagraphContainer>());
            }
        }

        private static void UpdateContentsWithParagraphs(SortedDictionary<decimal, List<ParagraphContainer>> documentContent, ParagraphContainer container)
        {
            if (container.Text.Length >= MinimalTextLength)
            {
                TryAddVerticalLocation(documentContent, container);
                documentContent[container.VerticalLocation] = InsertRangeInCollection(documentContent[container.VerticalLocation], container);
            }
        }

        private void GetDataFromPage(int i)
        {
            var pageContent = this.ReadSinglePage(i);
            this.UpdateLineMapping(pageContent);
        }

        private List<WordTable> GetTables()
        {
            List<WordTable> tables = new List<WordTable>(this.document.Tables.Count);

            for (int i = 1; i <= this.document.Tables.Count; i++)
            {
                tables.Add(new WordTable(this.document.Tables[i]));
            }

            return tables;
        }

        /// <summary>
        /// Gets TextFrame objects, which contain text, on specific document page.
        /// </summary>
        /// <param name="pageIndex">Index of page to read.</param>
        /// <returns>Collection of valid TextFrame objects.</returns>
        private List<TextFrame> GetValidTextFrames(long pageIndex)
        {
            var frames = new List<TextFrame>();
            for (int i = 1; i <= this.document.Shapes.Count; i++)
            {
                if (this.document.Shapes[i].Anchor.Information[WdInformation.wdActiveEndPageNumber] == pageIndex)
                {
                    TryAddFrame(frames, this.document.Shapes[i]);
                }
            }

            return frames;
        }

        /// <summary>
        /// Gets grouped document content from specified page.
        /// </summary>
        /// <param name="pageIndex">Index of page to read.</param>
        /// <returns>Document contents, grouped by lines.</returns>
        private SortedDictionary<long, List<ParagraphContainer>> ReadSinglePage(long pageIndex)
        {
            this.paragraphReader = this.paragraphReader ?? new WordParagraphReader(this.document, pageIndex);
            SortedDictionary<decimal, List<ParagraphContainer>> documentContent = this.paragraphReader.GetValidParagraphs(pageIndex);
            documentContent = this.UpdateContentsWithFrameContents(pageIndex, documentContent);
            return documentContent.Count == 0
                ? new SortedDictionary<long, List<ParagraphContainer>>()
                : GroupParagraphsByLine(documentContent);
        }

        private void TryAddTablesFromFrames(List<TextFrame> frames)
        {
            for (int frameIndex = 0; frameIndex < frames.Count; frameIndex++)
            {
                TextFrame item = frames[frameIndex];
                for (int i = 1; i <= item.TextRange.Tables.Count; i++)
                {
                    this.TableCollection.Add(new WordTable(item.TextRange.Tables[i]));
                }
            }
        }

        private SortedDictionary<decimal, List<ParagraphContainer>> UpdateContentsWithFrameContents(long pageIndex, SortedDictionary<decimal, List<ParagraphContainer>> documentContent)
        {
            List<TextFrame> frameCollection = this.GetValidTextFrames(pageIndex);
            this.TryAddTablesFromFrames(frameCollection);
            documentContent = AddDataFromFrames(documentContent, frameCollection);
            return documentContent;
        }

        /// <summary>
        /// Merges LineMapping dictionary with provided page content.
        /// </summary>
        /// <param name="pageContent">Grouped content from single page.</param>
        private void UpdateLineMapping(SortedDictionary<long, List<ParagraphContainer>> pageContent)
        {
            if (pageContent.Count == 0)
            {
                return;
            }

            List<long> keys = pageContent.Keys.ToList();
            if (pageContent.ContainsKey(0))
            {
                pageContent = ShiftContentKeys(pageContent, keys);
            }

            if (this.LineMapping.Count == 0)
            {
                this.LineMapping = pageContent;
                return;
            }

            this.UpdateLineMappingByEndLine(pageContent, keys);
        }

        private void UpdateLineMappingByEndLine(SortedDictionary<long, List<ParagraphContainer>> pageContent, List<long> keys)
        {
            long endLine = this.LineMapping.Keys.Last();
            for (int i = 0; i < keys.Count; i++)
            {
                long key = keys[i];
                this.LineMapping.Add(key + endLine, pageContent[key]);
            }
        }
    }
}