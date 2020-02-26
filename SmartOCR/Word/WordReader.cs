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
        /// Defines limits of single document line.
        /// 72 points = 1 inch.
        /// </summary>
        private const byte VerticalPositionOffset = 6;

        /// <summary>
        /// Defines lower bound of document text element length.
        /// </summary>
        private const byte MinimalTextLength = 2;

        /// <summary>
        /// Representation of Word document that is being read.
        /// </summary>
        private readonly Document document;

        /// <summary>
        /// Editable counter of read paragraphs.
        /// </summary>
        private int paragraphCounter = 1;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordReader"/> class.
        /// </summary>
        /// <param name="document">Word document representation.</param>
        public WordReader(Document document)
        {
            this.LineMapping = new SortedDictionary<long, List<ParagraphContainer>>();
            this.TableCollection = new List<WordTable>();
            this.document = document;
        }

        /// <summary>
        /// Gets document contents grouped in separate lines.
        /// </summary>
        public SortedDictionary<long, List<ParagraphContainer>> LineMapping { get; private set; }

        public List<WordTable> TableCollection { get; private set; }

        /// <summary>
        /// Performs data reading from all document pages.
        /// </summary>
        public void ReadDocument()
        {
            long numberOfPages = this.document.Range().Information[WdInformation.wdNumberOfPagesInDocument];

            for (int i = 1; i <= numberOfPages; i++)
            {
                var pageContent = this.ReadSinglePage(i);
                this.UpdateLineMapping(pageContent);
            }

            this.TableCollection = this.GetTables();
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
                ParagraphContainer container = paragraphContainers[i];
                decimal location = container.VerticalLocation;
                if (container.Text.Length >= MinimalTextLength)
                {
                    if (!documentContent.ContainsKey(location))
                    {
                        documentContent.Add(location, new List<ParagraphContainer>());
                    }

                    documentContent[location] = InsertRangeInCollection(documentContent[location], container);
                }
            }

            return documentContent;
        }

        /// <summary>
        /// Extracts paragraphs from <see cref="TextFrame"/> object.
        /// </summary>
        /// <param name="textFrame">A <see cref="TextFrame"/> instance that contains text.</param>
        /// <returns>A collection of <see cref="ParagraphContainer"/> objects.</returns>
        private static List<ParagraphContainer> GetParagraphsFromTextFrame(TextFrame textFrame)
        {
            Paragraphs paragraphs = textFrame.TextRange.Paragraphs;
            int paragraphsCount = paragraphs.Count;

            var paragraphContainers = new List<ParagraphContainer>();
            for (int i = 1; i <= paragraphsCount; i++)
            {
                paragraphContainers.Add(new ParagraphContainer(paragraphs[i].Range));
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
            var newDocumentContent = new SortedDictionary<long, List<ParagraphContainer>>();

            try
            {
                newDocumentContent.Add(1, documentContent.Values.First());
            }
            catch (InvalidOperationException)
            {
                return newDocumentContent;
            }

            for (int i = 1; i < documentContent.Count; i++)
            {
                decimal currentLocation = documentContent.Keys.ElementAt(i);
                decimal previousLocation = newDocumentContent[newDocumentContent.Count][0].VerticalLocation;
                if (previousLocation - VerticalPositionOffset <= currentLocation && currentLocation <= previousLocation + VerticalPositionOffset)
                {
                    newDocumentContent[newDocumentContent.Count].AddRange(documentContent[currentLocation]);
                }
                else
                {
                    newDocumentContent.Add(newDocumentContent.Count + 1, documentContent[currentLocation]);
                }
            }

            foreach (KeyValuePair<long, List<ParagraphContainer>> item in newDocumentContent)
            {
                item.Value.Sort();
            }

            return newDocumentContent;
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

        /// <summary>
        /// Gets data from Paragraph objects on specific page.
        /// </summary>
        /// <param name="pageIndex">Index of page to read.</param>
        /// <returns>Mapping of paragraphs, sorted by their vertical location on page.</returns>
        private SortedDictionary<decimal, List<ParagraphContainer>> GetDataFromParagraphs(long pageIndex)
        {
            var documentContent = new SortedDictionary<decimal, List<ParagraphContainer>>();
            List<ParagraphContainer> validParagraphs = this.GetValidParagraphs(pageIndex);

            for (int i = 0; i < validParagraphs.Count; i++)
            {
                ParagraphContainer singleParagraph = validParagraphs[i];
                decimal location = singleParagraph.VerticalLocation;
                if (!documentContent.ContainsKey(location))
                {
                    documentContent.Add(location, new List<ParagraphContainer>());
                }

                documentContent[location].Add(singleParagraph);
            }

            return documentContent;
        }

        private List<WordTable> GetTables()
        {
            List<WordTable> tables;
            if (this.TableCollection.Count == 0)
            {
                tables = new List<WordTable>(this.document.Tables.Count);
            }
            else
            {
                tables = this.TableCollection;
            }

            for (int i = 1; i <= this.document.Tables.Count; i++)
            {
                tables.Add(new WordTable(this.document.Tables[i]));
            }

            return tables;
        }

        /// <summary>
        /// Gets valid paragraphs, wrapped in <see cref="ParagraphContainer"/> instances, from specific page.
        /// </summary>
        /// <param name="pageIndex">Specific page index.</param>
        /// <returns>Collection of valid paragraphs.</returns>
        private List<ParagraphContainer> GetValidParagraphs(long pageIndex)
        {
            var paragraphs = this.document.Paragraphs;
            int paragraphsCount = paragraphs.Count;
            var paragraphCollection = new List<ParagraphContainer>();

            for (int i = this.paragraphCounter; i <= paragraphsCount; i++)
            {
                var singleRange = paragraphs[i].Range;
                long rangePage = singleRange.Information[WdInformation.wdActiveEndPageNumber];
                if (rangePage < pageIndex)
                {
                    continue;
                }

                if (rangePage > pageIndex)
                {
                    this.paragraphCounter = i;
                    break;
                }

                if (rangePage == pageIndex && singleRange.Text.Length > MinimalTextLength)
                {
                    paragraphCollection.Add(new ParagraphContainer(singleRange));
                }
            }

            return paragraphCollection;
        }

        /// <summary>
        /// Gets TextFrame objects, which contain text, on specific document page.
        /// </summary>
        /// <param name="pageIndex">Index of page to read.</param>
        /// <returns>Collection of valid TextFrame objects.</returns>
        private List<TextFrame> GetValidTextFrames(long pageIndex)
        {
            var frames = new List<TextFrame>();
            var shapes = this.document.Shapes;
            int shapesCount = shapes.Count;
            for (int i = 1; i <= shapesCount; i++)
            {
                var shape = shapes[i];
                long shapePage = shape.Anchor.Information[WdInformation.wdActiveEndPageNumber];
                if (shapePage < pageIndex)
                {
                    continue;
                }

                if (shapePage > pageIndex)
                {
                    break;
                }

                var frame = shape.TextFrame;
                if (frame != null && frame.HasText != 0 && frame.TextRange.Text.Length > MinimalTextLength)
                {
                    frames.Add(frame);
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
            SortedDictionary<decimal, List<ParagraphContainer>> documentContent = this.GetDataFromParagraphs(pageIndex);
            List<TextFrame> frameCollection = this.GetValidTextFrames(pageIndex);
            this.TryAddTablesFromFrames(frameCollection);
            documentContent = AddDataFromFrames(documentContent, frameCollection);
            return GroupParagraphsByLine(documentContent);
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
                var shiftedMapping = new SortedDictionary<long, List<ParagraphContainer>>();
                for (int i = 0; i < keys.Count; i++)
                {
                    shiftedMapping.Add(keys[i] + 1, pageContent[keys[i]]);
                }

                pageContent = shiftedMapping;
            }

            if (this.LineMapping.Count == 0)
            {
                this.LineMapping = pageContent;
                return;
            }

            long endLine = this.LineMapping.Keys.Last();
            for (int i = 0; i < keys.Count; i++)
            {
                long key = keys[i];
                this.LineMapping.Add(key + endLine, pageContent[key]);
            }
        }
    }
}