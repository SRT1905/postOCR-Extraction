namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to test document line contents for matching with regular expression.
    /// </summary>
    public class LineContentChecker
    {
        /// <summary>
        /// Collection of paragraphs to check.
        /// </summary>
        private readonly List<ParagraphContainer> paragraphs = new List<ParagraphContainer>();

        /// <summary>
        /// Indicates whether all paragraphs or some selection should be tested.
        /// </summary>
        private readonly int searchStatus;

        /// <summary>
        /// Indicates index of last paragraph to test.
        /// </summary>
        private int finishIndex;

        /// <summary>
        /// Indicates index of first paragraph to test.
        /// </summary>
        private int startIndex;

        /// <summary>
        /// Initializes a new instance of the <see cref="LineContentChecker"/> class.
        /// Instance has collection of paragraphs to test.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        public LineContentChecker(List<ParagraphContainer> paragraphs)
        {
            this.paragraphs = paragraphs ?? throw new ArgumentNullException(nameof(paragraphs));
            this.startIndex = 0;
            this.finishIndex = paragraphs.Count - 1;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LineContentChecker"/> class.
        /// Instance has collection of paragraphs to test.
        /// Passed location and search status define selection of paragraphs.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        /// <param name="paragraphLocation">Location of margin paragraph.</param>
        /// <param name="searchStatus">Indicates whether test is performed on all paragraphs or by some margin. Must be in range [-1; 1].</param>
        /// <exception cref="ArgumentOutOfRangeException">Invalid <paramref name="searchStatus"/> value is provided.</exception>
        public LineContentChecker(List<ParagraphContainer> paragraphs, decimal paragraphLocation, int searchStatus)
            : this(paragraphs)
        {
            this.ParagraphHorizontalLocation = paragraphLocation;
            if (Math.Abs(searchStatus) > 1)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(searchStatus), Properties.Resources.outOfRangeParagraphHorizontalLocationStatus);
            }

            this.searchStatus = searchStatus;
            this.SetSearchIndexes();
        }

        /// <summary>
        /// Gets or sets all matches in paragraph.
        /// </summary>
        public string JoinedMatches { get; set; }

        /// <summary>
        /// Gets or sets position of first/last paragraph to test and matched paragraph location.
        /// </summary>
        public decimal ParagraphHorizontalLocation { get; set; }

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

            for (int location = this.startIndex; location <= this.finishIndex; location++)
            {
                string paragraphText = this.paragraphs[location].Text;
                if (regExObject.IsMatch(paragraphText))
                {
                    if (string.IsNullOrEmpty(checkValue))
                    {
                        var foundMatches = GetMatchesFromParagraph(regExObject, paragraphText);
                        this.JoinedMatches = string.Join("|", foundMatches);
                        this.ParagraphHorizontalLocation = this.paragraphs[location].HorizontalLocation;
                        return true;
                    }
                    else
                    {
                        var foundMatches = GetMatchesFromParagraph(regExObject, paragraphText, checkValue);
                        if (foundMatches.Count != 0)
                        {
                            this.JoinedMatches = string.Join("|", foundMatches.Select(item => item.Value));
                            this.ParagraphHorizontalLocation = this.paragraphs[location].HorizontalLocation;
                            return true;
                        }
                    }
                }
            }

            this.ParagraphHorizontalLocation = 0;
            return false;
        }

        /// <summary>
        /// Gets matches between regular expression and text.
        /// </summary>
        /// <param name="regexObject"><see cref="Regex"/> object that is used to test text.</param>
        /// <param name="textToCheck">Text to check.</param>
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
        /// <param name="textToCheck">Text to check.</param>
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

        /// <summary>
        /// Looks up index of paragraph, which horizontal location matches passed location.
        /// </summary>
        /// <param name="returnNextLargest">Indicates that, if location is not matched, index of paragraph with next biggest (or next smallest) location is returned.</param>
        /// <returns>Index of paragraph with matching location.</returns>
        private int GetParagraphByLocation(bool returnNextLargest)
        {
            List<int> locations = this.paragraphs.Select(item => (int)item.HorizontalLocation).ToList();
            int location = locations.BinarySearch((int)this.ParagraphHorizontalLocation);
            if (location < 0)
            {
                location = ~location;
            }

            if (returnNextLargest)
            {
                if (location == this.paragraphs.Count)
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
            switch (this.searchStatus)
            {
                case 1:
                    this.startIndex = this.GetParagraphByLocation(true);
                    this.finishIndex = this.paragraphs.Count - 1;
                    break;
                case -1:
                    this.startIndex = 0;
                    this.finishIndex = this.GetParagraphByLocation(false);
                    break;
                default:
                    break;
            }
        }
    }
}
