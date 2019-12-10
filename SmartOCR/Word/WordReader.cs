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
            return document_content;
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

        protected SortedDictionary<long, List<ParagraphContainer>> AddDataFromFrames(SortedDictionary<long, List<ParagraphContainer>> document_content, List<TextFrame> frame_collection)
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

        protected SortedDictionary<long, List<ParagraphContainer>> GetDataFromParagraphs(long page_index)
        {
            SortedDictionary<long, List<ParagraphContainer>> document_content = new SortedDictionary<long, List<ParagraphContainer>>();
            for (int i = 1; i <= document.Paragraphs.Count; i++)
            {
                Range single_range = document.Paragraphs[i].Range;
                if (single_range.Information[WdInformation.wdActiveEndPageNumber] > page_index)
                {
                    return document_content;
                }
                if (single_range.Text.Length <= minimal_text_length || single_range.Information[WdInformation.wdActiveEndPageNumber] < page_index)
                {
                    continue;
                }

                long line_number = single_range.Information[WdInformation.wdFirstCharacterLineNumber];
                if (!document_content.ContainsKey(line_number))
                {
                    document_content.Add(line_number, new List<ParagraphContainer>());
                }

                document_content[line_number].Add(new ParagraphContainer(single_range));
            }
            return document_content;
        }

        private SortedDictionary<long, List<ParagraphContainer>> AddDataFromSingleFrame(SortedDictionary<long, List<ParagraphContainer>> document_content, TextFrame text_frame)
        {
            foreach (Paragraph paragraph in text_frame.TextRange.Paragraphs)
            {
                ParagraphContainer container = new ParagraphContainer(paragraph.Range);
                if (container.Text.Length > minimal_text_length)
                {
                    bool add_new_line = false;
                    long closest_line = GetClosestLine(document_content, container.VerticalLocation, ref add_new_line);
                    if (!document_content.ContainsKey(closest_line))
                    {
                        document_content.Add(closest_line, new List<ParagraphContainer>());
                    }
                    else if (add_new_line)
                    {
                        document_content = ShiftReadLines(document_content, closest_line);
                    }

                    document_content[closest_line] = InsertRangeInCollection(document_content[closest_line], container);
                }
            }
            return document_content;
        }

        private long GetClosestLine(SortedDictionary<long, List<ParagraphContainer>> document_content, double position, ref bool add_new_line)
        {
            var verticals = document_content.Values.Select(item => item.First().VerticalLocation).ToList();
            int line = verticals.BinarySearch(position);
            if (line < 0)
            {
                line = ~line;
                return GetApproximateLineNumber(document_content, position, ref add_new_line, line - 1, line);
            }
            if (line == 0)
            {
                return GetApproximateLineNumber(document_content, position, ref add_new_line, 0, 0);
            }
            if (~line == verticals.Count)
            {
                line = ~line;
                return GetApproximateLineNumber(document_content, position, ref add_new_line, line - 1, line - 1);
            }
            return line;
        }

        private long GetApproximateLineNumber(SortedDictionary<long, List<ParagraphContainer>> document_content, double position, ref bool add_new_line, int lower_index, int upper_index)
        {
            List<long> keys = document_content.Keys.ToList();
            long line;
            double line_position;
            if (lower_index == 0)
            {
                line = keys[lower_index];
                line_position = document_content[line][0].VerticalLocation;
                return AdjustLineNumber(position, 0, 0, line, line_position, ref add_new_line);
            }
            if (upper_index == document_content.Count - 1)
            {
                line = keys[upper_index];
                line_position = document_content[line][0].VerticalLocation;
                return AdjustLineNumber(position, line, line_position, 10000, 10000, ref add_new_line);
            }
            line = keys[lower_index];
            line_position = document_content[line][0].VerticalLocation;
            double adjacent_position;
            long adjacent_line;
            if (line_position <= position)
            {
                adjacent_line = keys[lower_index + 1];
                adjacent_position = document_content[adjacent_line][0].VerticalLocation;
                return AdjustLineNumber(position, line, line_position, adjacent_line, adjacent_position, ref add_new_line);
            }
            adjacent_line = keys[lower_index - 1];
            adjacent_position = document_content[adjacent_line][0].VerticalLocation;
            return AdjustLineNumber(position, adjacent_line, adjacent_position, line, line_position, ref add_new_line);
        }

        private long AdjustLineNumber(double position, long upper_line, double upper_line_position, long lower_line, double lower_line_position, ref bool add_new_line)
        {
            int threshold = lower_line_position - upper_line_position <= shape_position_offset
                ? 0
                : shape_position_offset;
            bool inside_upper_position = upper_line_position <= position
                                      && position <= upper_line_position + threshold;
            bool inside_lower_position = lower_line_position - threshold <= position
                                      && position <= lower_line_position;
            bool between_positions = !(inside_lower_position || inside_upper_position);

            add_new_line = false;
            if (between_positions)
            {
                add_new_line = true;
                return position - upper_line_position > lower_line_position - position
                    ? lower_line - 1
                    : upper_line + 1;
            }
            if (inside_upper_position)
            {
                return upper_line;
            }
            if (inside_lower_position)
            {
                return lower_line;
            }
            return 0;
        }

        private SortedDictionary<long, List<ParagraphContainer>> ShiftReadLines(SortedDictionary<long, List<ParagraphContainer>> document_content, long line)
        {
            SortedDictionary<long, List<ParagraphContainer>> dict = new SortedDictionary<long, List<ParagraphContainer>>();
            long shift = 0;
            foreach (long current_key in document_content.Keys)
            {
                if (current_key == line)
                {
                    dict.Add(current_key, new List<ParagraphContainer>());
                    shift = 1;
                }
                dict.Add(current_key + shift, document_content[current_key]);
            }
            return dict;
        }

        private List<ParagraphContainer> InsertRangeInCollection(List<ParagraphContainer> paragraph_collection, ParagraphContainer text_range_container)
        {
            paragraph_collection.Add(text_range_container);
            paragraph_collection.Sort();
            return paragraph_collection;
        }
    }
}