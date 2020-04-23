namespace SmartOCR.Word.Read
{
    using System.Collections.Generic;
    using Microsoft.Office.Interop.Word;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

    /// <summary>
    /// Used to collect data from Word text frames.
    /// </summary>
    public class WordTextFrameReader : ITableReader
    {
        /// <summary>
        /// Defines lower bound of document text element length.
        /// </summary>
        private const byte MinimalTextLength = 2;
        private readonly Document document;
        private int pageIndex;

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
        public List<WordTable> GetWordTables(int pageIndex)
        {
            this.pageIndex = pageIndex;
            return this.GetWordTables();
        }

        /// <inheritdoc/>
        public List<WordTable> GetWordTables()
        {
            var tables = new List<WordTable>(this.document.Tables.Count);
            this.CollectTablesFromFrames(tables);
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
            PopulateParagraphMapping(frameCollection, documentContent);

            return documentContent;
        }

        private static void PopulateParagraphMapping(List<TextFrame> frameCollection, ParagraphMapping documentContent)
        {
            foreach (var frame in frameCollection)
            {
                AddDataFromSingleFrame(documentContent, frame);
            }
        }

        /// <summary>
        /// Gets data from TextFrame object and adds it to document contents representation.
        /// </summary>
        /// <param name="documentContent">Representation of read document contents.</param>
        /// <param name="textFrame">TextFrame object containing text.</param>
        private static void AddDataFromSingleFrame(ParagraphMapping documentContent, TextFrame textFrame)
        {
            List<ParagraphContainer> paragraphContainers = GetParagraphsFromTextFrame(textFrame);
            foreach (var paragraph in paragraphContainers)
            {
                AddParagraphToContents(documentContent, paragraph);
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
            for (int frameParagraphIndex = 1; frameParagraphIndex <= textFrame.TextRange.Paragraphs.Count; frameParagraphIndex++)
            {
                TryAddFrameContent(textFrame, paragraphContainers, frameParagraphIndex);
            }

            return paragraphContainers;
        }

        private static void TryAddFrameContent(TextFrame textFrame, List<ParagraphContainer> paragraphContainers, int frameParagraphIndex)
        {
            if (textFrame.TextRange.Paragraphs[frameParagraphIndex].Range.Text.Length > MinimalTextLength)
            {
                paragraphContainers.Add(new ParagraphContainer(textFrame.TextRange.Paragraphs[frameParagraphIndex].Range));
            }
        }

        private static void AddParagraphToContents(ParagraphMapping documentContent, ParagraphContainer paragraph)
        {
            decimal location = paragraph.VerticalLocation;
            TryAddNewLocationToDocumentContent(documentContent, location);

            documentContent[location].Add(paragraph);
        }

        private static void TryAddNewLocationToDocumentContent(ParagraphMapping documentContent, decimal location)
        {
            if (!documentContent.ContainsKey(location))
            {
                documentContent.Add(location, new List<ParagraphContainer>());
            }
        }

        private void CollectTablesFromFrames(List<WordTable> tables)
        {
            var frames = this.GetValidTextFrames();

            foreach (var frame in frames)
            {
                GetTablesFromSingleFrame(tables, frame);
            }
        }

        private List<TextFrame> GetValidTextFrames()
        {
            var frames = new List<TextFrame>();
            for (int shapeIndex = 1; shapeIndex <= this.document.Shapes.Count; shapeIndex++)
            {
                this.ExtractFramesFromShapes(frames, shapeIndex);
            }

            return frames;
        }

        private void ExtractFramesFromShapes(List<TextFrame> frames, int shapeIndex)
        {
            if (this.document.Shapes[shapeIndex]
                             .Anchor
                             .Information[WdInformation.wdActiveEndPageNumber] == this.pageIndex)
            {
                TryAddFrame(frames, this.document.Shapes[shapeIndex]);
            }
        }
    }
}
