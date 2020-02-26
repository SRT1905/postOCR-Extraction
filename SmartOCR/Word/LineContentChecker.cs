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
        private int finishIndex;
        /// <summary>
        /// Indicates whether all paragraphs or some selection should be tested.
        /// </summary>
        private readonly int searchStatus;
        /// <summary>
        /// Indicates index of first paragraph to test.
        /// </summary>
        private int startIndex;
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
            startIndex = 0;
            finishIndex = paragraphs.Count - 1;
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
            this.searchStatus = searchStatus;
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
            for (int location = startIndex; location <= finishIndex; location++)
            {
                string paragraphText = paragraphs[location].Text;
                if (regExObject.IsMatch(paragraphText))
                {
                    if (string.IsNullOrEmpty(checkValue))
                    {
                        var foundMatches = GetMatchesFromParagraph(regExObject, paragraphText);
                        JoinedMatches = string.Join("|", foundMatches);
                        ParagraphHorizontalLocation = paragraphs[location].HorizontalLocation;
                        return true;
                    }
                    else
                    {
                        var foundMatches = GetMatchesFromParagraph(regExObject, paragraphText, checkValue);
                        if (foundMatches.Count != 0)
                        {
                            JoinedMatches = string.Join("|", foundMatches.Select(item => item.Value));
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
        /// <param name="returnNextLargest">Indicates that, if location is not matched, 
        /// index of paragraph with next biggest (or next smallest) location is returned.</param>
        /// <returns>Index of paragraph with matching location.</returns>
        private int GetParagraphByLocation(bool returnNextLargest)
        {
            List<int> locations = paragraphs.Select(item => (int)item.HorizontalLocation).ToList();
            int location = locations.BinarySearch((int)ParagraphHorizontalLocation);
            if (location < 0)
            {
                location = ~location;
            }
            if (returnNextLargest)
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
            switch (searchStatus)
            {
                case 1:
                    startIndex = GetParagraphByLocation(true);
                    finishIndex = paragraphs.Count - 1;
                    break;
                case -1:
                    startIndex = 0;
                    finishIndex = GetParagraphByLocation(false);
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
        /// <param name="regexObject"><see cref="Regex"/> object that is used to test text.</param>
        /// <param name="textToCheck"></param>
        /// <returns>Collection of found matches.</returns>
        private static List<string> GetMatchesFromParagraph(Regex regexObject, string textToCheck)
        {
            MatchCollection matches = regexObject.Matches(textToCheck);
            var foundValues = new List<string>();
            for (int i = 0; i < matches.Count; i++)
            {
                Match singleMatch = matches[i];
                if (singleMatch.Groups.Count > 1)
                {
                    for (int groupIndex = 1; groupIndex < singleMatch.Groups.Count; groupIndex++)
                    {
                        Group groupItem = singleMatch.Groups[groupIndex];
                        foundValues.Add(groupItem.Value);

                    }
                }
                else
                {
                    foundValues.Add(singleMatch.Value);
                }
            }
            return foundValues;
        }
        /// <summary>
        /// Gets collection of <see cref="SimilarityDescription"/> objects which indicate ratio of similarity between matched text and check value.
        /// </summary>
        /// <param name="regexObject"><see cref="Regex"/> object that is used to test text.</param>
        /// <param name="textToCheck"></param>
        /// <param name="checkValue">Value used for similarity check with found match.</param>
        /// <returns>Collection of <see cref="SimilarityDescription"/> objects, containing matches and similarity ratios.</returns>
        private static List<SimilarityDescription> GetMatchesFromParagraph(Regex regexObject, string textToCheck, string checkValue)
        {
            MatchCollection matches = regexObject.Matches(textToCheck);
            List<SimilarityDescription> foundValues = new List<SimilarityDescription>();
            for (int i = 0; i < matches.Count; i++)
            {
                Match singleMatch = matches[i];
                if (singleMatch.Groups.Count > 1)
                {
                    for (int groupIndex = 1; groupIndex < singleMatch.Groups.Count; groupIndex++)
                    {
                        Group groupItem = singleMatch.Groups[groupIndex];
                        SimilarityDescription description = new SimilarityDescription(groupItem.Value, checkValue);
                        if (description.CheckStringSimilarity())
                        {
                            foundValues.Add(description);
                        }
                    }
                }
                else
                {
                    SimilarityDescription description = new SimilarityDescription(singleMatch.Value, checkValue);
                    if (description.CheckStringSimilarity())
                    {
                        foundValues.Add(description);
                    }
                }
            }
            return foundValues;
        }
        #endregion
    }
}
