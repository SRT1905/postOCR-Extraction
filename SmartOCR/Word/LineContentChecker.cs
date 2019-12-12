using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    internal class LineContentChecker
    {
        private List<ParagraphContainer> Paragraphs { get; }
        public decimal paragraph_horizontal_location;
        public string joined_matches;
        private readonly int search_status;
        private int start_index;
        private int finish_index;

        public LineContentChecker()
        {
            this.Paragraphs = new List<ParagraphContainer>();
        }

        public LineContentChecker(List<ParagraphContainer> paragraphs)
        {
            this.Paragraphs = paragraphs;
        }

        public LineContentChecker(List<ParagraphContainer> paragraphs, decimal paragraph_index, int search_status) : this(paragraphs)
        {
            this.paragraph_horizontal_location = paragraph_index;
            this.search_status = search_status;
            SetSearchIndexes();
        }

        private void SetSearchIndexes()
        {
            switch (search_status)
            {
                case 0:
                    start_index = 0;
                    finish_index = Paragraphs.Count - 1;
                    break;
                case 1:
                    start_index = GetParagraphByLocation(true);
                    finish_index = Paragraphs.Count - 1;
                    break;
                case -1:
                    start_index = 0;
                    finish_index = GetParagraphByLocation(false);
                    break;
                default:
                    break;
            }
        }

        private int GetParagraphByLocation(bool return_next_largest)
        {
            List<double> locations = Paragraphs.Select(item => item.HorizontalLocation).ToList();
            int location = locations.BinarySearch(paragraph_horizontal_location);
            if (location < 0)
            { 
                location = ~location;
            }
            if (return_next_largest)
            {
                if (location == Paragraphs.Count)
                {
                    return location--;
                }
                return location;
            }
            return location--;
        }


        public bool CheckLineContents(Regex regex_obj, string check_value)
        {
            for (int location = start_index; location <= finish_index; location++)
            {
                string paragraph_text = Paragraphs[location].Text;
                if (regex_obj.IsMatch(paragraph_text))
                {
                    if (string.IsNullOrEmpty(check_value))
                    {
                        var found_matches = GetMatchesFromParagraph(paragraph_text, regex_obj);
                        this.joined_matches = string.Join("|", found_matches);
                        paragraph_horizontal_location = Paragraphs[location].HorizontalLocation;
                        return true;
                    }
                    else
                    {
                        var found_matches = GetMatchesFromParagraph(paragraph_text, regex_obj, check_value);
                        if (found_matches.Count != 0)
                        {
                            this.joined_matches = string.Join("|", found_matches.Select(item => item.Value));
                            paragraph_horizontal_location = Paragraphs[location].HorizontalLocation;
                            return true;
                        }
                    }

                }
            }
            paragraph_horizontal_location = 0;
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