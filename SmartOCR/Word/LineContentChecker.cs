using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    /// <summary>
    /// Used to test document line contents for matching with regular expression.
    /// </summary>
    public class LineContentChecker
    {
        #region Fields
        /// <summary>
        /// Indicates index of last paragraph to test.
        /// </summary>
        private int finish_index;
        /// <summary>
        /// Indicates whether all paragraphs or some selection should be tested.
        /// </summary>
        private readonly int search_status;
        /// <summary>
        /// Indicates index of first paragraph to test.
        /// </summary>
        private int start_index;
        /// <summary>
        /// Collection of paragraphs to check.
        /// </summary>
        private readonly List<ParagraphContainer> paragraphs = new List<ParagraphContainer>();
        #endregion

        #region Properties
        /// <summary>
        /// Represents all matches in paragraph.
        /// </summary>
        public string JoinedMatches { get; set; }
        /// <summary>
        /// Used to indicate position of first/last paragraph to test and to indicate matched paragraph location.
        /// </summary>
        public decimal ParagraphHorizontalLocation { get; set; }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes instance of <see cref="LineContentChecker"/> that has collection of paragraphs to test.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        public LineContentChecker(List<ParagraphContainer> paragraphs)
        {
            this.paragraphs = paragraphs ?? throw new ArgumentNullException(nameof(paragraphs));
            start_index = 0;
            finish_index = paragraphs.Count - 1;
        }
        /// <summary>
        /// Initializes instance of <see cref="LineContentChecker"/> that has collection of paragraphs to test.
        /// Passed location and search status define selection of paragraphs.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        /// <param name="paragraphLocation">Location of margin paragraph.</param>
        /// <param name="searchStatus">Indicates whether test is performed on all paragraphs or by some margin. Must be in range [-1; 1].</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public LineContentChecker(List<ParagraphContainer> paragraphs, decimal paragraphLocation, int searchStatus) : this(paragraphs)
        {
            this.ParagraphHorizontalLocation = paragraphLocation;
            if (Math.Abs(searchStatus) > 1)
            {
                throw new ArgumentOutOfRangeException(nameof(searchStatus),
                                                      Properties.Resources.outOfRangeParagraphHorizontalLocationStatus);
            }
            this.search_status = searchStatus;
            SetSearchIndexes();
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Tests paragraphs for matching with regular expression and performs similarity check with passed value.
        /// </summary>
        /// <param name="regExObject"><see cref="Regex"/> object that is used to test paragraphs.</param>
        /// <param name="checkValue">Value used for similarity check with found match. Can be null, then similarity check is not performed.</param>
        /// <returns>Indicator whether there is match between regular expression and paragraph contents.</returns>
        public bool CheckLineContents(Regex regExObject, string checkValue)
        {
            if (regExObject == null)
            {
                throw new ArgumentNullException(nameof(regExObject));
            }
            for (int location = start_index; location <= finish_index; location++)
            {
                string paragraph_text = paragraphs[location].Text;
                if (regExObject.IsMatch(paragraph_text))
                {
                    if (string.IsNullOrEmpty(checkValue))
                    {
                        var found_matches = GetMatchesFromParagraph(regExObject, paragraph_text);
                        this.JoinedMatches = string.Join("|", found_matches);
                        ParagraphHorizontalLocation = paragraphs[location].HorizontalLocation;
                        return true;
                    }
                    else
                    {
                        var found_matches = GetMatchesFromParagraph(regExObject, paragraph_text, checkValue);
                        if (found_matches.Count != 0)
                        {
                            this.JoinedMatches = string.Join("|", found_matches.Select(item => item.Value));
                            ParagraphHorizontalLocation = paragraphs[location].HorizontalLocation;
                            return true;
                        }
                    }

                }
            }
            ParagraphHorizontalLocation = 0;
            return false;
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Looks up index of paragraph, which horizontal location matches passed location.
        /// </summary>
        /// <param name="return_next_largest">Indicates that, if location is not matched, 
        /// index of paragraph with next biggest (or next smallest) location is returned.</param>
        /// <returns>Index of paragraph with matching location.</returns>
        private int GetParagraphByLocation(bool return_next_largest)
        {
            List<int> locations = paragraphs.Select(item => (int)item.HorizontalLocation).ToList();
            int location = locations.BinarySearch((int)ParagraphHorizontalLocation);
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
        #endregion

        #region Private static methods
        /// <summary>
        /// Gets matches between regular expression and text.
        /// </summary>
        /// <param name="re_object"><see cref="Regex"/> object that is used to test text.</param>
        /// <param name="text_to_check"></param>
        /// <returns>Collection of found matches.</returns>
        private static List<string> GetMatchesFromParagraph(Regex re_object, string text_to_check)
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
        private static List<SimilarityDescription> GetMatchesFromParagraph(Regex re_object, string text_to_check, string check_value)
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
        #endregion
    }
}
