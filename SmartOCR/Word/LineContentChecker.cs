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
        private readonly List<ParagraphContainer> paragraphs = new List<ParagraphContainer>();
        private readonly int searchStatus;
        private int finishIndex;
        private int startIndex = 0;

        /// <summary>
        /// Initializes a new instance of the <see cref="LineContentChecker"/> class.
        /// Instance has collection of paragraphs to test.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        public LineContentChecker(List<ParagraphContainer> paragraphs)
        {
            this.paragraphs = paragraphs ?? throw new ArgumentNullException(nameof(paragraphs));
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
            ValidateSearchStatus(searchStatus);

            this.ParagraphHorizontalLocation = paragraphLocation;
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
            for (int location = this.startIndex; location <= this.finishIndex; location++)
            {
                if (regExObject.IsMatch(this.paragraphs[location].Text))
                {
                    if (string.IsNullOrEmpty(checkValue))
                    {
                        return this.ProcessMatchWithoutCheck(regExObject, location, this.paragraphs[location].Text);
                    }

                    if (this.ProcessMatchWithCheck(regExObject, location, checkValue))
                    {
                        return true;
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
                GetValuesFromParagraphSingleMatch(matches, foundValues, i);
            }

            return foundValues;
        }

        private static void GetValuesFromParagraphSingleMatch(MatchCollection matches, List<string> foundValues, int i)
        {
            if (matches[i].Groups.Count > 1)
            {
                for (int groupIndex = 1; groupIndex < matches[i].Groups.Count; groupIndex++)
                {
                    foundValues.Add(matches[i].Groups[groupIndex].Value);
                }
            }
            else
            {
                foundValues.Add(matches[i].Value);
            }
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
                GetValuesFromParagraphSingleMatch(matches, foundValues, i, checkValue);
            }

            return foundValues;
        }

        private static void GetValuesFromParagraphSingleMatch(MatchCollection matches, List<SimilarityDescription> foundValues, int i, string checkValue)
        {
            Match singleMatch = matches[i];
            if (singleMatch.Groups.Count > 1)
            {
                ProcessMultipleGroupsWithCheck(foundValues, checkValue, singleMatch);
            }
            else
            {
                ProcessSingleMatchWithCheck(foundValues, checkValue, singleMatch);
            }
        }

        private static void ProcessSingleMatchWithCheck(List<SimilarityDescription> foundValues, string checkValue, Match singleMatch)
        {
            SimilarityDescription description = new SimilarityDescription(singleMatch.Value, checkValue);
            if (description.AreStringsSimilar())
            {
                foundValues.Add(description);
            }
        }

        private static void ProcessMultipleGroupsWithCheck(List<SimilarityDescription> foundValues, string checkValue, Match singleMatch)
        {
            for (int groupIndex = 1; groupIndex < singleMatch.Groups.Count; groupIndex++)
            {
                Group groupItem = singleMatch.Groups[groupIndex];
                SimilarityDescription description = new SimilarityDescription(groupItem.Value, checkValue);
                if (description.AreStringsSimilar())
                {
                    foundValues.Add(description);
                }
            }
        }

        private static void ValidateSearchStatus(int searchStatus)
        {
            if (Math.Abs(searchStatus) > 1)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(searchStatus), Properties.Resources.outOfRangeParagraphHorizontalLocationStatus);
            }
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

        private bool ProcessMatchWithCheck(Regex regExObject, int location, string checkValue)
        {
            var foundMatches = GetMatchesFromParagraph(regExObject, this.paragraphs[location].Text, checkValue);
            if (foundMatches.Count == 0)
            {
                return false;
            }

            this.JoinedMatches = string.Join("|", foundMatches.Select(item => item.Value));
            this.ParagraphHorizontalLocation = this.paragraphs[location].HorizontalLocation;
            return true;
        }

        private bool ProcessMatchWithoutCheck(Regex regExObject, int location, string paragraphText)
        {
            var foundMatches = GetMatchesFromParagraph(regExObject, paragraphText);
            this.JoinedMatches = string.Join("|", foundMatches);
            this.ParagraphHorizontalLocation = this.paragraphs[location].HorizontalLocation;
            return true;
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
