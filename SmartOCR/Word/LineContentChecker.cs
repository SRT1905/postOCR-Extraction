using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    /// <summary>
    /// Used to test document line contents for matching with regular expression.
    /// </summary>
    internal class LineContentChecker
    {
        /// <summary>
        /// Collection of paragraphs to check.
        /// </summary>
        private readonly List<ParagraphContainer> paragraphs = new List<ParagraphContainer>();
   
        /// <summary>
        /// Indicates whether all paragraphs or some selection should be tested.
        /// </summary>
        private readonly int search_status;
 
        /// <summary>
        /// Indicates index of first paragraph to test.
        /// </summary>
        private int start_index;
  
        /// <summary>
        /// Indicates index of last paragraph to test.
        /// </summary>
        private int finish_index;
 
        /// <summary>
        /// Used to indicate position of first/last paragraph to test and to indicate matched paragraph location.
        /// </summary>
        public decimal paragraph_horizontal_location;
 
        /// <summary>
        /// Represents all matches in paragraph.
        /// </summary>
        public string joined_matches;

        /// <summary>
        /// Initializes instance of <see cref="LineContentChecker"/> that has collection of paragraphs to test.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        public LineContentChecker(List<ParagraphContainer> paragraphs)
        {
            this.paragraphs = paragraphs;
            start_index = 0;
            finish_index = paragraphs.Count - 1;
        }

        /// <summary>
        /// Initializes instance of <see cref="LineContentChecker"/> that has collection of paragraphs to test.
        /// Passed location and search status define selection of paragraphs.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        /// <param name="paragraph_location">Location of margin paragraph.</param>
        /// <param name="search_status">Indicates whether test is performed on all paragraphs or by some margin. Must be in range [-1; 1].</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public LineContentChecker(List<ParagraphContainer> paragraphs, decimal paragraph_location, int search_status) : this(paragraphs)
        {
            this.paragraph_horizontal_location = paragraph_location;
            if (Math.Abs(search_status) > 1)
            {
                throw new ArgumentOutOfRangeException(nameof(search_status), Properties.Resources.outOfRangeParagraphHorizontalLocationStatus);
            }
            this.search_status = search_status;
            SetSearchIndexes();
        }

        /// <summary>
        /// Defines lower and upper bounds of testing paragraphs.
        /// </summary>
        private void SetSearchIndexes()
        {
            switch (search_status)
            {
                case 1:
                    start_index = GetParagraphByLocation(true);
                    finish_index = paragraphs.Count - 1;
                    break;
                case -1:
                    start_index = 0;
                    finish_index = GetParagraphByLocation(false);
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Looks up index of paragraph, which horizontal location matches passed location.
        /// </summary>
        /// <param name="return_next_largest">Indicates that, if location is not matched, 
        /// index of paragraph with next biggest (or next smallest) location is returned.</param>
        /// <returns>Index of paragraph with matching location.</returns>
        private int GetParagraphByLocation(bool return_next_largest)
        {
            List<int> locations = paragraphs.Select(item => (int)item.HorizontalLocation).ToList();
            int location = locations.BinarySearch((int)paragraph_horizontal_location);
            if (location < 0)
            { 
                location = ~location;
            }
            if (return_next_largest)
            {
                if (location == paragraphs.Count)
                {
                    return location--;
                }
                return location;
            }
            return location--;
        }

        /// <summary>
        /// Tests paragraphs for matching with regular expression and performs similarity check with passed value.
        /// </summary>
        /// <param name="regex_obj"><see cref="Regex"/> object that is used to test paragraphs.</param>
        /// <param name="check_value">Value used for similarity check with found match. Can be null, then similarity check is not performed.</param>
        /// <returns>Indicator whether there is match between regular expression and paragraph contents.</returns>
        public bool CheckLineContents(Regex regex_obj, string check_value)
        {
            for (int location = start_index; location <= finish_index; location++)
            {
                string paragraph_text = paragraphs[location].Text;
                if (regex_obj.IsMatch(paragraph_text))
                {
                    if (string.IsNullOrEmpty(check_value))
                    {
                        var found_matches = GetMatchesFromParagraph(regex_obj, paragraph_text);
                        this.joined_matches = string.Join("|", found_matches);
                        paragraph_horizontal_location = paragraphs[location].HorizontalLocation;
                        return true;
                    }
                    else
                    {
                        var found_matches = GetMatchesFromParagraph(regex_obj, paragraph_text, check_value);
                        if (found_matches.Count != 0)
                        {
                            this.joined_matches = string.Join("|", found_matches.Select(item => item.Value));
                            paragraph_horizontal_location = paragraphs[location].HorizontalLocation;
                            return true;
                        }
                    }

                }
            }
            paragraph_horizontal_location = 0;
            return false;
        }

        /// <summary>
        /// Gets matches between regular expression and text.
        /// </summary>
        /// <param name="re_object"><see cref="Regex"/> object that is used to test text.</param>
        /// <param name="text_to_check"></param>
        /// <returns>Collection of found matches.</returns>
        private List<string> GetMatchesFromParagraph(Regex re_object, string text_to_check)
        {
            MatchCollection matches = re_object.Matches(text_to_check);
            var found_values = new List<string>();
            for (int i = 0; i < matches.Count; i++)
            {
                Match single_match = matches[i];
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

        /// <summary>
        /// Gets collection of <see cref="SimilarityDescription"/> objects which indicate ratio of similarity between matched text and check value.
        /// </summary>
        /// <param name="re_object"><see cref="Regex"/> object that is used to test text.</param>
        /// <param name="text_to_check"></param>
        /// <param name="check_value">Value used for similarity check with found match.</param>
        /// <returns>Collection of <see cref="SimilarityDescription"/> objects, containing matches and similarity ratios.</returns>
        private List<SimilarityDescription> GetMatchesFromParagraph(Regex re_object, string text_to_check, string check_value)
        {
            MatchCollection matches = re_object.Matches(text_to_check);
            List<SimilarityDescription> found_values = new List<SimilarityDescription>();
            for (int i = 0; i < matches.Count; i++)
            {
                Match single_match = matches[i];
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
