using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    internal class WordReader
    {
        public SortedDictionary<long, List<ParagraphContainer>> line_mapping;
        private readonly byte shape_position_offset = 5;
        private readonly byte minimal_text_length = 2;

        public WordReader(Document document)
        {
            line_mapping = ReadDocument(document);
            WordApplication.CloseDocument(document);
        }

        protected SortedDictionary<long, List<ParagraphContainer>> ReadDocument(Document document)
        {
            var document_content = GetDataFromParagraphs(document);
            List<TextFrame> frame_collection = GetValidTextFrames(document);
            AddDataFromFrames(ref document_content, frame_collection);
            return document_content;
        }

        protected void AddDataFromFrames(ref SortedDictionary<long, List<ParagraphContainer>> document_content, List<TextFrame> frame_collection)
        {
            foreach (TextFrame frame in frame_collection)
            {
                if (frame.TextRange.Text.Length > minimal_text_length)
                {
                    AddDataFromSingleFrame(ref document_content, frame);
                }
            }
        }

        protected List<TextFrame> GetValidTextFrames(Document document)
        {
            List<TextFrame> frame_collection = new List<TextFrame>();
            for (int i = 1; i <= document.Shapes.Count; i++)
            {
                TextFrame frame = document.Shapes[i].TextFrame;
                if (frame != null && frame.HasText != 0)
                {
                    frame_collection.Add(frame);
                }
            }
            return frame_collection;
        }

        protected SortedDictionary<long, List<ParagraphContainer>> GetDataFromParagraphs(Document document)
        {
            SortedDictionary<long, List<ParagraphContainer>> document_content = new SortedDictionary<long, List<ParagraphContainer>>();
            for (int i = 1; i <= document.Paragraphs.Count; i++)
            {
                Range single_range = document.Paragraphs[i].Range;
                if (single_range.Text.Length <= minimal_text_length)
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

        private void AddDataFromSingleFrame(ref SortedDictionary<long, List<ParagraphContainer>> document_content, TextFrame text_frame)
        {
            foreach (Paragraph paragraph in text_frame.TextRange.Paragraphs)
            {
                ParagraphContainer container = new ParagraphContainer(paragraph.Range);
                if (container.Text.Length > minimal_text_length)
                {
                    bool add_new_line = false;
                    long closest_line = GetClosestLine(ref document_content, container.VerticalLocation, ref add_new_line);
                    if (!document_content.ContainsKey(closest_line))
                    {
                        document_content.Add(closest_line, new List<ParagraphContainer>());
                    }
                    else if (add_new_line)
                    {
                        document_content = ShiftReadLines(ref document_content, closest_line);
                    }

                    InsertRangeInCollection(ref document_content, closest_line, container);
                }
            }
        }

        private long GetClosestLine(ref SortedDictionary<long, List<ParagraphContainer>> document_content, double position, ref bool add_new_line)
        {
            int lower_index = 0;
            int upper_index = document_content.Count - 1;
            while (lower_index <= upper_index)
            {
                int middle_index = lower_index + ((upper_index - lower_index) / 2);
                List<long> keys = document_content.Keys.ToList();
                long middle_line = keys[middle_index];
                double middle_position = document_content[middle_line][0].VerticalLocation;
                if (position == middle_position)
                {
                    long adjacent_line = middle_index - 1 < 0 ? keys[0] : keys[middle_index - 1];
                    double adjacent_position = document_content[adjacent_line][0].VerticalLocation;
                    return AdjustLineNumber(position, adjacent_line, adjacent_position, middle_line, middle_position, ref add_new_line);
                }
                if (lower_index == upper_index)
                {
                    return GetApproximateLineNumber(ref document_content, position, ref add_new_line, lower_index, upper_index);
                }
                if (position > middle_position)
                {
                    lower_index = middle_index + 1;
                }
                else if (position < middle_position)
                {
                    upper_index = middle_index - 1;
                }
            }
            return 0;
        }

        private long GetApproximateLineNumber(ref SortedDictionary<long, List<ParagraphContainer>> document_content, double position, ref bool add_new_line, int lower_index, int upper_index)
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

        private SortedDictionary<long, List<ParagraphContainer>> ShiftReadLines(ref SortedDictionary<long, List<ParagraphContainer>> document_content, long line)
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

        private void InsertRangeInCollection(ref SortedDictionary<long, List<ParagraphContainer>> document_content, long closest_line, ParagraphContainer text_range_container)
        {
            document_content[closest_line].Add(text_range_container);
            document_content[closest_line].Sort();
        }
    }
}