using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    internal class WordReader : IDisposable
    {
        public SortedDictionary<long, List<ParagraphContainer>> line_mapping;
        public Document document;
        private const byte shape_position_offset = 5;
        private const byte minimal_text_length = 2;

        public WordReader()
        {
            line_mapping = new SortedDictionary<long, List<ParagraphContainer>>();
        }
        public WordReader(Document document) : this()
        {
            this.document = document;
        }

        public void Dispose()
        {
            if (document != null)
            {
                WordApplication.CloseDocument(document);
            }
        }

        public void ReadDocument()
        {
            long number_of_pages = document.Range().Information[WdInformation.wdNumberOfPagesInDocument];

            for (int i = 1; i <= number_of_pages; i++)
            {
                var page_content = ReadSinglePage(i);
                UpdateLineMapping(page_content);
            }
        }

        public void ReadDocument(long page_index)
        {
            line_mapping = ReadSinglePage(page_index);
        }

        private SortedDictionary<long, List<ParagraphContainer>> ReadSinglePage(long page_index)
        {
            var document_content = GetDataFromParagraphs(page_index);
            List<TextFrame> frame_collection = GetValidTextFrames(page_index);
            document_content = AddDataFromFrames(document_content, frame_collection);
            return GroupParagraphsByLine(document_content);
        }

        private void UpdateLineMapping(SortedDictionary<long, List<ParagraphContainer>> page_content)
        {
            if (page_content.Count == 0)
            {
                return;
            }

            if (page_content.ContainsKey(0))
            {
                var shifted_mapping = new SortedDictionary<long, List<ParagraphContainer>>();
                foreach (long key in page_content.Keys)
                {
                    shifted_mapping.Add(key + 1, page_content[key]);
                }
                page_content = shifted_mapping;
            }

            if (line_mapping.Count == 0)
            {
                line_mapping = page_content;
                return;
            }

            long end_line = line_mapping.Keys.Last();
            foreach (long key in page_content.Keys)
            {
                line_mapping.Add(key + end_line, page_content[key]);
            }


        }

        protected SortedDictionary<decimal, List<ParagraphContainer>> AddDataFromFrames(SortedDictionary<decimal, List<ParagraphContainer>> document_content, List<TextFrame> frame_collection)
        {
            foreach (TextFrame frame in frame_collection)
            {
                if (frame.TextRange.Text.Length > minimal_text_length)
                {
                    document_content = AddDataFromSingleFrame(document_content, frame);
                }
            }
            return document_content;
        }

        protected List<TextFrame> GetValidTextFrames(long page_index)
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

        protected SortedDictionary<decimal, List<ParagraphContainer>> GetDataFromParagraphs(long page_index)
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

            foreach (ParagraphContainer item in valid_paragraphs)
            {
                if (!document_content.ContainsKey(item.VerticalLocation))
                {
                    document_content.Add(item.VerticalLocation, new List<ParagraphContainer>());
                }
                document_content[item.VerticalLocation].Add(item);
            }
            return document_content;
        }

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

                if (previous_location - shape_position_offset <= current_location && current_location <= previous_location + 5)
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

        private SortedDictionary<decimal, List<ParagraphContainer>> AddDataFromSingleFrame(SortedDictionary<decimal, List<ParagraphContainer>> document_content, TextFrame text_frame)
        {
            List<ParagraphContainer> paragraph_containers = new List<ParagraphContainer>();
            foreach (Paragraph item in text_frame.TextRange.Paragraphs)
            {
                paragraph_containers.Add(new ParagraphContainer(item.Range));
            }
            foreach (ParagraphContainer container in paragraph_containers)
            {
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

        private List<ParagraphContainer> InsertRangeInCollection(List<ParagraphContainer> paragraph_collection, ParagraphContainer text_range_container)
        {
            paragraph_collection.Add(text_range_container);
            paragraph_collection.Sort();
            return paragraph_collection;
        }
    }
}