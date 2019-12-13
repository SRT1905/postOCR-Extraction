using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// Performs collection and grouping of document contents.
    /// </summary>
    internal class WordReader : IDisposable
    {
        /// <summary>
        /// Defines limits of single document line.
        /// </summary>
        private const byte vertical_position_offset = 5;
       
        /// <summary>
        /// Defines lower bound of document text element length.
        /// </summary>
        private const byte minimal_text_length = 2;
        
        /// <summary>
        /// Representation of Word document that is being read.
        /// </summary>
        private readonly Document document;
        
        /// <summary>
        /// Represents document contents grouped in separate lines.
        /// </summary>
        public SortedDictionary<long, List<ParagraphContainer>> LineMapping { get; set; }

        /// <summary>
        /// Initializes instance of WordReader class that has document to read.
        /// </summary>
        /// <param name="document">Word document representation.</param>
        public WordReader(Document document)
        {
            LineMapping = new SortedDictionary<long, List<ParagraphContainer>>();
            this.document = document;
        }

        /// <summary>
        /// Closes Word document if it is open.
        /// </summary>
        public void Dispose()
        {
            if (document != null)
            {
                WordApplication.CloseDocument(document);
            }
        }

        /// <summary>
        /// Performs data reading from all document pages.
        /// </summary>
        public void ReadDocument()
        {
            long number_of_pages = document.Range().Information[WdInformation.wdNumberOfPagesInDocument];

            for (int i = 1; i <= number_of_pages; i++)
            {
                var page_content = ReadSinglePage(i);
                UpdateLineMapping(page_content);
            }
        }

        /// <summary>
        /// Performs data reading from single document page.
        /// </summary>
        /// <param name="page_index">Index of page to read, starting from 1.</param>
        public void ReadDocument(long page_index)
        {
            if (page_index >= 1)
            {
                LineMapping = ReadSinglePage(page_index);
            }
        }

        /// <summary>
        /// Gets grouped document content from specified page.
        /// </summary>
        /// <param name="page_index">Index of page to read.</param>
        /// <returns>Document contents, grouped by lines.</returns>
        private SortedDictionary<long, List<ParagraphContainer>> ReadSinglePage(long page_index)
        {
            var document_content = GetDataFromParagraphs(page_index);
            List<TextFrame> frame_collection = GetValidTextFrames(page_index);
            document_content = AddDataFromFrames(document_content, frame_collection);
            return GroupParagraphsByLine(document_content);
        }

        /// <summary>
        /// Merges LineMapping dictionary with provided page content.
        /// </summary>
        /// <param name="page_content">Grouped content from single page.</param>
        private void UpdateLineMapping(SortedDictionary<long, List<ParagraphContainer>> page_content)
        {
            if (page_content.Count == 0)
            {
                return;
            }
            List<long> keys = page_content.Keys.ToList();
            if (page_content.ContainsKey(0))
            {
                var shifted_mapping = new SortedDictionary<long, List<ParagraphContainer>>();
                for (int i = 0; i < keys.Count; i++)
                {
                    long key = keys[i];
                    shifted_mapping.Add(key + 1, page_content[key]);
                }
                page_content = shifted_mapping;
            }
            if (LineMapping.Count == 0)
            {
                LineMapping = page_content;
                return;
            }
            long end_line = LineMapping.Keys.Last();
            for (int i = 0; i < keys.Count; i++)
            {
                long key = keys[i];
                LineMapping.Add(key + end_line, page_content[key]);
            }
        }

        /// <summary>
        /// Gets data from Paragraph objects on specific page.
        /// </summary>
        /// <param name="page_index">Index of page to read.</param>
        /// <returns>Mapping of paragraphs, sorted by their vertical location on page.</returns>
        private SortedDictionary<decimal, List<ParagraphContainer>> GetDataFromParagraphs(long page_index)
        {
            var document_content = new SortedDictionary<decimal, List<ParagraphContainer>>();
            var valid_paragraphs = new List<ParagraphContainer>();

            for (int i = 1; i <= document.Paragraphs.Count; i++)
            {
                Range range = document.Paragraphs[i].Range;
                int page_number = range.Information[WdInformation.wdActiveEndPageNumber];
                if (page_number < page_index)
                {
                    continue;
                }
                if (page_number > page_index)
                {
                    break;
                }
                if (range.Text.Length > minimal_text_length)
                {
                    valid_paragraphs.Add(new ParagraphContainer(range));
                }
            }

            for (int i = 0; i < valid_paragraphs.Count; i++)
            {
                ParagraphContainer single_paragraph = valid_paragraphs[i];
                decimal location = single_paragraph.VerticalLocation;
                if (!document_content.ContainsKey(location))
                {
                    document_content.Add(location, new List<ParagraphContainer>());
                }
                document_content[location].Add(single_paragraph);
            }
            return document_content;
        }

        /// <summary>
        /// Gets TextFrame objects, which contain text, on specific document page.
        /// </summary>
        /// <param name="page_index">Index of page to read.</param>
        /// <returns>Collection of valid TextFrame objects.</returns>
        private List<TextFrame> GetValidTextFrames(long page_index)
        {
            List<TextFrame> frame_collection = new List<TextFrame>();
            for (int i = 1; i <= document.Shapes.Count; i++)
            {
                Shape shape = document.Shapes[i];
                if (shape.Anchor.Information[WdInformation.wdActiveEndPageNumber] > page_index)
                {
                    return frame_collection;
                }
                if (shape.Anchor.Information[WdInformation.wdActiveEndPageNumber] < page_index)
                {
                    continue;
                }
                TextFrame frame = shape.TextFrame;
                if (frame != null && frame.HasText != 0)
                {
                    frame_collection.Add(frame);
                }
            }
            return frame_collection;
        }

        /// <summary>
        /// Gets data from TextFrame objects and adds it to document contents container.
        /// </summary>
        /// <param name="document_content">Representation of read document contents.</param>
        /// <param name="frame_collection">Collection of TextFrame objects.</param>
        /// <returns>Representation of document contents that is extended by TextFrame objects.</returns>
        private SortedDictionary<decimal, List<ParagraphContainer>> AddDataFromFrames(SortedDictionary<decimal, List<ParagraphContainer>> document_content, List<TextFrame> frame_collection)
        {
            for (int i = 0; i < frame_collection.Count; i++)
            {
                TextFrame frame = frame_collection[i];
                if (frame.TextRange.Text.Length > minimal_text_length)
                {
                    document_content = AddDataFromSingleFrame(document_content, frame);
                }
            }
            return document_content;
        }

        /// <summary>
        /// Gets data from TextFrame object and adds it to document contents representation.
        /// </summary>
        /// <param name="document_content">Representation of read document contents.</param>
        /// <param name="text_frame">TextFrame object containing text.</param>
        /// <returns>Representation of read document contents, extended by TextFrame contents.</returns>
        private SortedDictionary<decimal, List<ParagraphContainer>> AddDataFromSingleFrame(SortedDictionary<decimal, List<ParagraphContainer>> document_content, TextFrame text_frame)
        {
            List<ParagraphContainer> paragraph_containers = new List<ParagraphContainer>();
            foreach (Paragraph item in text_frame.TextRange.Paragraphs)
            {
                paragraph_containers.Add(new ParagraphContainer(item.Range));
            }
            for (int i = 0; i < paragraph_containers.Count; i++)
            {
                ParagraphContainer container = paragraph_containers[i];
                decimal location = container.VerticalLocation;
                if (container.Text.Length >= minimal_text_length)
                {
                    if (!document_content.ContainsKey(location))
                    {
                        document_content.Add(location, new List<ParagraphContainer>());
                    }
                    document_content[location] = InsertRangeInCollection(document_content[location], container);
                }
            }
            return document_content;
        }
        
        /// <summary>
        /// Groups document contents, arranged by vertical location on page, in separate lines.
        /// </summary>
        /// <param name="document_content">Representation of read document contents.</param>
        /// <returns></returns>
        private SortedDictionary<long, List<ParagraphContainer>> GroupParagraphsByLine(SortedDictionary<decimal, List<ParagraphContainer>> document_content)
        {
            var new_document_content = new SortedDictionary<long, List<ParagraphContainer>>()
            {
                {1, document_content.Values.First()},
            };
            for (int i = 1; i < document_content.Count; i++)
            {
                decimal current_location = document_content.Keys.ElementAt(i);
                decimal previous_location = document_content.Keys.ElementAt(i - 1);
                if (previous_location - vertical_position_offset <= current_location && current_location <= previous_location + 5)
                {
                    if (new_document_content.ContainsKey(i))
                    {
                        new_document_content[i].AddRange(document_content[current_location]);
                        new_document_content[i].Sort();
                    }
                    else
                    {
                        new_document_content.Add(i + 1, document_content[current_location]);
                    }
                }
                else
                {
                    new_document_content.Add(i + 1, document_content[current_location]);
                }
            }
            return new_document_content;
        }
        
        /// <summary>
        /// Adds ParagraphContainer instance to collection with maintaining sort order.
        /// </summary>
        /// <param name="paragraph_collection">Collection of ParagraphContainer objects.</param>
        /// <param name="text_range_container">ParagraphContainer instance to add.</param>
        /// <returns>Updated collection of ParagraphContainer objects.</returns>
        private List<ParagraphContainer> InsertRangeInCollection(List<ParagraphContainer> paragraph_collection, ParagraphContainer text_range_container)
        {
            paragraph_collection.Add(text_range_container);
            paragraph_collection.Sort();
            return paragraph_collection;
        }
    }
}