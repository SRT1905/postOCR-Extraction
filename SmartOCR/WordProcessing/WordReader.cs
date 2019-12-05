using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    class WordReader
    {
        const byte shape_position_offset = 5;
        public SortedDictionary<long, List<ParagraphContainer>> line_mapping;

        public WordReader(Document document)
        {
            line_mapping = ReadDocument(document);
        }

        private SortedDictionary<long, List<ParagraphContainer>> ReadDocument(Document document)
        {
            SortedDictionary<long, List<ParagraphContainer>> line_mapping = new SortedDictionary<long, List<ParagraphContainer>>();
            for (int i = 1; i <= document.Paragraphs.Count; i++)
            {
                Range single_range = document.Paragraphs[i].Range;
                if (single_range.Text.Length <= 2)
                {
                    continue;
                }
                long line_number = single_range.Information[WdInformation.wdFirstCharacterLineNumber];
                if (!line_mapping.ContainsKey(line_number))
                {
                    line_mapping.Add(line_number, new List<ParagraphContainer>());
                }

                line_mapping[line_number].Add(new ParagraphContainer(single_range));
            }
            List<TextFrame> frame_collection = new List<TextFrame>();
            for (int i = 1; i <= document.Shapes.Count; i++)
            {
                frame_collection.Add(document.Shapes[i].TextFrame);
            }

            foreach (TextFrame frame in frame_collection)
            {
                if (frame != null && frame.HasText != 0)
                {
                    if (frame.TextRange.Text.Length > 2)
                    {
                        AddDataFromShape(ref line_mapping, frame);
                    }
                }
            }
            return line_mapping;
        }
        private void AddDataFromShape(ref SortedDictionary<long, List<ParagraphContainer>> line_mapping, TextFrame text_frame)
        {
            foreach (Paragraph paragraph in text_frame.TextRange.Paragraphs)
            {
                ParagraphContainer container = new ParagraphContainer(paragraph.Range);
                if (container.Text.Length > 2)
                {
                    bool add_new_line = false;
                    long closest_line = GetClosestLine(line_mapping, container.VerticalLocation, ref add_new_line);
                    if (!line_mapping.ContainsKey(closest_line))
                    {
                        line_mapping.Add(closest_line, new List<ParagraphContainer>());
                    }
                    else if (add_new_line)
                    {
                        line_mapping = ShiftReadLines(line_mapping, closest_line);
                    }

                    InsertRangeInCollection(ref line_mapping, closest_line, container);
                }
            }
        }
        private long GetClosestLine(SortedDictionary<long, List<ParagraphContainer>> line_mapping, double position, ref bool add_new_line)
        {
            int lower_index = 0;
            int upper_index = line_mapping.Count - 1;
            List<long> keys = line_mapping.Keys.ToList();

            while (lower_index <= upper_index)
            {
                int middle_index = lower_index + ((upper_index - lower_index) / 2);
                long middle_line = keys[middle_index];
                double middle_position = line_mapping[middle_line][0].VerticalLocation;
                long adjacent_line;
                double adjacent_position;
                if (position == middle_position)
                {
                    adjacent_line = middle_index - 1 < 0 ? keys[0] : keys[middle_index - 1];
                    adjacent_position = line_mapping[adjacent_line][0].VerticalLocation;
                    return AdjustLineNumber(position, adjacent_line, adjacent_position, middle_line, middle_position, ref add_new_line);
                }
                if (lower_index == upper_index)
                {
                    long line;
                    double line_position;
                    if (lower_index == 0)
                    {
                        line = keys[lower_index];
                        line_position = line_mapping[line][0].VerticalLocation;
                        return AdjustLineNumber(position, 0, 0, line, line_position, ref add_new_line);
                    }
                    if (upper_index == line_mapping.Count - 1)
                    {
                        line = keys[upper_index];
                        line_position = line_mapping[line][0].VerticalLocation;
                        return AdjustLineNumber(position, line, line_position, 10000, 10000, ref add_new_line);
                    }
                    line = keys[lower_index];
                    line_position = line_mapping[line][0].VerticalLocation;
                    if (line_position <= position)
                    {
                        adjacent_line = keys[lower_index + 1];
                        adjacent_position = line_mapping[adjacent_line][0].VerticalLocation;
                        return AdjustLineNumber(position, line, line_position, adjacent_line, adjacent_position, ref add_new_line);
                    }
                    adjacent_line = keys[lower_index - 1];
                    adjacent_position = line_mapping[adjacent_line][0].VerticalLocation;
                    return AdjustLineNumber(position, adjacent_line, adjacent_position, line, line_position, ref add_new_line);
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
        private long AdjustLineNumber(double position, long upper_line, double upper_line_position, long lower_line, double lower_line_position, ref bool add_new_line)
        {

            int threshold = lower_line_position - upper_line_position <= shape_position_offset ? 0 : shape_position_offset;
            bool inside_upper_position = upper_line_position <= position && position <= upper_line_position + threshold;
            bool inside_lower_position = lower_line_position - threshold <= position && position <= lower_line_position;
            bool between_positions = !(inside_lower_position || inside_upper_position);

            add_new_line = false;
            if (between_positions)
            {
                add_new_line = true;
                return position - upper_line_position > lower_line_position - position ? lower_line - 1 : upper_line + 1;
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
        private SortedDictionary<long, List<ParagraphContainer>> ShiftReadLines(SortedDictionary<long, List<ParagraphContainer>> line_mapping, long line)
        {
            SortedDictionary<long, List<ParagraphContainer>> dict = new SortedDictionary<long, List<ParagraphContainer>>();
            long shift = 0;
            for (int i = 0; i < line_mapping.Count; i++)
            {
                long current_key = line_mapping.Keys.ToList()[i];
                if (current_key == line)
                {
                    dict.Add(current_key, new List<ParagraphContainer>());
                    shift = 1;
                }
                dict.Add(current_key + shift, line_mapping[current_key]);
            }
            return dict;
        }
        private void InsertRangeInCollection(ref SortedDictionary<long, List<ParagraphContainer>> line_mapping, long closest_line, ParagraphContainer text_range_container)
        {
            List<ParagraphContainer> paragraph_collection = line_mapping[closest_line];
            int i;
            for (i = 0; i < paragraph_collection.Count; i++)
            {
                ParagraphContainer container = paragraph_collection[i];
                if (container.HorizontalLocation > text_range_container.HorizontalLocation)
                    break;
            }
            if (i == paragraph_collection.Count)
            {
                paragraph_collection.Add(text_range_container);
            }
            else
            {
                paragraph_collection.Insert(i, text_range_container);
            }
            line_mapping[closest_line] = paragraph_collection;
        }
        public void Dispose()
        {
            WordApplication.CloseActiveWordDocument();
        }
    }
}
