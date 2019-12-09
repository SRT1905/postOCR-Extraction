using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    internal class LineContentChecker
    {
        private List<ParagraphContainer> paragraph_collection { get; }
        public long paragraph_index;
        public string joined_matches;

        public LineContentChecker()
        {
            this.paragraph_collection = new List<ParagraphContainer>();
        }

        public LineContentChecker(List<ParagraphContainer> paragraphs)
        {
            this.paragraph_collection = paragraphs;
        }

        public LineContentChecker(List<ParagraphContainer> paragraphs, long paragraph_index) : this(paragraphs)
        {
            this.paragraph_index = paragraph_index;
        }

        public bool CheckLineContents(Regex regex_obj, string check_value)
        {
            long start_index = paragraph_index == 0 || paragraph_index >= paragraph_collection.Count ? 0 : paragraph_index;

            for (paragraph_index = start_index; paragraph_index < paragraph_collection.Count; paragraph_index++)
            {
                string paragraph_text = paragraph_collection[(int)paragraph_index].Text;
                if (regex_obj.IsMatch(paragraph_text))
                {
                    if (string.IsNullOrEmpty(check_value))
                    {
                        var found_matches = GetMatchesFromParagraph(paragraph_text, regex_obj);
                        this.joined_matches = string.Join("|", found_matches);
                        return true;
                    }
                    else
                    {
                        var found_matches = GetMatchesFromParagraph(paragraph_text, regex_obj, check_value);
                        if (found_matches.Count != 0)
                        {
                            this.joined_matches = string.Join("|", found_matches.Select(item => item.Value));
                            return true;
                        }
                    }

                }
            }
            paragraph_index = 0;
            return false;
        }
        private List<string> GetMatchesFromParagraph(string text_to_check, Regex re_object)
        {
            MatchCollection matches = re_object.Matches(text_to_check);
            var found_values = new List<string>();
            foreach (Match single_match in matches)
            {
                if (single_match.Groups.Count > 1)
                {
                    for (int group_i = 1; group_i < single_match.Groups.Count; group_i++)
                    {
                        Group group_item = single_match.Groups[group_i];
                        found_values.Add(group_item.Value);
                        
                    }
                }
                else
                {
                    found_values.Add(single_match.Value);
                }
            }
            return found_values;
        }

        private List<SimilarityDescription> GetMatchesFromParagraph(string text_to_check, Regex re_object, string check_value)
        {
            MatchCollection matches = re_object.Matches(text_to_check);
            List<SimilarityDescription> found_values = new List<SimilarityDescription>();
            foreach (Match single_match in matches)
            {
                if (single_match.Groups.Count > 1)
                {
                    for (int group_i = 1; group_i < single_match.Groups.Count; group_i++)
                    {
                        Group group_item = single_match.Groups[group_i];
                        SimilarityDescription description = new SimilarityDescription(group_item.Value, check_value);
                        if (description.CheckStringSimilarity())
                        {
                            found_values.Add(description);
                        }
                    }
                }
                else
                {
                    SimilarityDescription description = new SimilarityDescription(single_match.Value, check_value);
                    if (description.CheckStringSimilarity())
                    {
                        found_values.Add(description);
                    }
                }
            }
            return found_values;
        }
    }
}