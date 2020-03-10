namespace SmartOCR
{
    using System.Collections.Generic;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Used to collect data from Word text frames.
    /// </summary>
    internal class WordTextFrameReader : ITableReader
    {
        /// <summary>
        /// Defines lower bound of document text element length.
        /// </summary>
        private const byte MinimalTextLength = 2;
        private readonly Document document;
        private readonly int pageIndex;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTextFrameReader"/> class.
        /// </summary>
        /// <param name="document">A Word document.</param>
        /// <param name="pageIndex">Page index, that contains text frames to read.</param>
        public WordTextFrameReader(Document document, int pageIndex)
        {
            this.document = document;
            this.pageIndex = pageIndex;
        }

        /// <summary>
        /// Returns mapping between paragraph horizontal location and collection of paragraphs at location.
        /// </summary>
        /// <returns>A mapping between paragraph horizontal location and collection of paragraphs at location.</returns>
        public ParagraphMapping GetParagraphMapping()
        {
            Utilities.Debug($"Getting text in shapes from page {this.pageIndex}.", 3);
            return AddDataFromFrames(this.GetValidTextFrames());
        }

        /// <inheritdoc/>
        public List<WordTable> GetWordTables()
        {
            List<WordTable> tables = new List<WordTable>(this.document.Tables.Count);
            var frames = this.GetValidTextFrames();

            for (int frameIndex = 0; frameIndex < frames.Count; frameIndex++)
            {
                GetTablesFromSingleFrame(tables, frames[frameIndex]);
            }

            return tables;
        }

        private static void GetTablesFromSingleFrame(List<WordTable> tables, TextFrame frame)
        {
            for (int tableIndex = 1; tableIndex <= frame.TextRange.Tables.Count; tableIndex++)
            {
                tables.Add(new WordTable(frame.TextRange.Tables[tableIndex]));
            }
        }

        private static void TryAddFrame(List<TextFrame> frames, Shape shape)
        {
            if (shape.TextFrame != null &&
                shape.TextFrame.HasText != 0 &&
                shape.TextFrame.TextRange.Text.Length > MinimalTextLength)
            {
                frames.Add(shape.TextFrame);
            }
        }

        /// <summary>
        /// Gets data from TextFrame objects and adds it to document contents container.
        /// </summary>
        /// <param name="frameCollection">Collection of TextFrame objects.</param>
        /// <returns>Representation of document contents that is extended by TextFrame objects.</returns>
        private static ParagraphMapping AddDataFromFrames(List<TextFrame> frameCollection)
        {
            var documentContent = new ParagraphMapping();
            for (int i = 0; i < frameCollection.Count; i++)
            {
                AddDataFromSingleFrame(documentContent, frameCollection[i]);
            }

            return documentContent;
        }

        /// <summary>
        /// Gets data from TextFrame object and adds it to document contents representation.
        /// </summary>
        /// <param name="documentContent">Representation of read document contents.</param>
        /// <param name="textFrame">TextFrame object containing text.</param>
        private static void AddDataFromSingleFrame(ParagraphMapping documentContent, TextFrame textFrame)
        {
            List<ParagraphContainer> paragraphContainers = GetParagraphsFromTextFrame(textFrame);
            for (int paragraphIndex = 0; paragraphIndex < paragraphContainers.Count; paragraphIndex++)
            {
                AddParagraphToContents(documentContent, paragraphContainers[paragraphIndex]);
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
                if (textFrame.TextRange.Paragraphs[i].Range.Text.Length > MinimalTextLength)
                {
                    paragraphContainers.Add(new ParagraphContainer(textFrame.TextRange.Paragraphs[i].Range));
                }
            }

            return paragraphContainers;
        }

        private static void AddParagraphToContents(ParagraphMapping documentContent, ParagraphContainer paragraph)
        {
            decimal location = paragraph.VerticalLocation;
            if (!documentContent.ContainsKey(location))
            {
                documentContent.Add(location, new List<ParagraphContainer>());
            }

            documentContent[location].Add(paragraph);
        }

        private List<TextFrame> GetValidTextFrames()
        {
            var frames = new List<TextFrame>();
            for (int i = 1; i <= this.document.Shapes.Count; i++)
            {
                if (this.document.Shapes[i]
                                 .Anchor
                                 .Information[WdInformation.wdActiveEndPageNumber] == this.pageIndex)
                {
                    TryAddFrame(frames, this.document.Shapes[i]);
                }
            }

            return frames;
        }
    }
}
