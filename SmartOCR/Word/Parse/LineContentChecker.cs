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
        private readonly bool useSoundex;
        private int finishIndex;
        private int startIndex = 0;

        /// <summary>
        /// Initializes a new instance of the <see cref="LineContentChecker"/> class.
        /// Instance has collection of paragraphs to test.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        /// <param name="useSoundex">Indicates whether <see cref="ParagraphContainer.Soundex"/> property should be used instead of <see cref="ParagraphContainer.Text"/>.</param>
        public LineContentChecker(List<ParagraphContainer> paragraphs, bool useSoundex)
        {
            this.paragraphs = paragraphs;
            this.finishIndex = paragraphs.Count - 1;
            this.useSoundex = useSoundex;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LineContentChecker"/> class.
        /// Instance has collection of paragraphs to test.
        /// Passed location and search status define selection of paragraphs.
        /// </summary>
        /// <param name="paragraphs">Collection of paragraphs.</param>
        /// <param name="useSoundex">Indicates whether <see cref="ParagraphContainer.Soundex"/> property should be used instead of <see cref="ParagraphContainer.Text"/>.</param>
        /// <param name="paragraphLocation">Location of margin paragraph.</param>
        /// <param name="searchStatus">Indicates whether test is performed on all paragraphs or by some margin. Must be in range [-1; 1].</param>
        /// <exception cref="ArgumentOutOfRangeException">Invalid <paramref name="searchStatus"/> value is provided.</exception>
        public LineContentChecker(List<ParagraphContainer> paragraphs, bool useSoundex, decimal paragraphLocation, int searchStatus)
            : this(paragraphs, useSoundex)
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
                string paragraphContent = this.GetTextToCheckByItsLocation(location);
                if (regExObject.IsMatch(paragraphContent))
                {
                    if (string.IsNullOrEmpty(checkValue))
                    {
                        return this.ProcessMatchWithoutCheck(regExObject, location, paragraphContent);
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
                GetValuesFromParagraphSingleMatch(matches[i], foundValues);
            }

            return foundValues;
        }

        private static void GetValuesFromParagraphSingleMatch(Match match, List<string> foundValues)
        {
            if (match.Groups.Count > 1)
            {
                GetValuesFromParagraphGroupsMatch(match, foundValues);
            }
            else
            {
                foundValues.Add(match.Value);
            }
        }

        private static void GetValuesFromParagraphGroupsMatch(Match match, List<string> foundValues)
        {
            for (int groupIndex = 1; groupIndex < match.Groups.Count; groupIndex++)
            {
                foundValues.Add(match.Groups[groupIndex].Value);
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
            foreach (Match match in matches)
            {
                GetValuesFromParagraphSingleMatch(match, foundValues, checkValue);
            }

            return foundValues;
        }

        private static void GetValuesFromParagraphSingleMatch(Match singleMatch, List<SimilarityDescription> foundValues, string checkValue)
        {
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
                TryAddSimilarityDescription(foundValues, checkValue, singleMatch.Groups[groupIndex]);
            }
        }

        private static void TryAddSimilarityDescription(List<SimilarityDescription> foundValues, string checkValue, Group singleGroup)
        {
            SimilarityDescription description = new SimilarityDescription(singleGroup.Value, checkValue);
            if (description.AreStringsSimilar())
            {
                foundValues.Add(description);
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

        private static int ValidateNegativeParagraphLocation(int location)
        {
            return location < 0
                ? ~location
                : location;
        }

        private string GetTextToCheckByItsLocation(int location)
        {
            return this.useSoundex
                                ? this.paragraphs[location].Soundex
                                : this.paragraphs[location].Text;
        }

        /// <summary>
        /// Looks up index of paragraph, which horizontal location matches passed location.
        /// </summary>
        /// <param name="returnNextLargest">Indicates that, if location is not matched, index of paragraph with next biggest (or next smallest) location is returned.</param>
        /// <returns>Index of paragraph with matching location.</returns>
        private int GetParagraphByLocation(bool returnNextLargest)
        {
            int location = this.paragraphs.Select(item => (int)item.HorizontalLocation)
                                          .ToList()
                                          .BinarySearch((int)this.ParagraphHorizontalLocation);
            return this.GetValidatedParagraphLocation(
                returnNextLargest,
                ValidateNegativeParagraphLocation(location));
        }

        private int GetValidatedParagraphLocation(bool returnNextLargest, int location)
        {
            return !returnNextLargest || location == this.paragraphs.Count
                ? --location
                : location;
        }

        private bool ProcessMatchWithCheck(Regex regExObject, int location, string checkValue)
        {
            List<SimilarityDescription> foundMatches = this.GetMatchesWithCheckValue(regExObject, location, checkValue);
            this.CombineMatchesAndSetFoundLocation(location, foundMatches);
            return foundMatches.Count != 0;
        }

        private void CombineMatchesAndSetFoundLocation(int location, List<SimilarityDescription> foundMatches)
        {
            if (foundMatches.Count != 0)
            {
                this.JoinedMatches = string.Join("|", foundMatches.Select(item => item.Value));
                this.ParagraphHorizontalLocation = this.paragraphs[location].HorizontalLocation;
            }
        }

        private List<SimilarityDescription> GetMatchesWithCheckValue(Regex regExObject, int location, string checkValue)
        {
            return this.useSoundex
                ? GetMatchesFromParagraph(regExObject, this.paragraphs[location].Soundex, checkValue)
                : GetMatchesFromParagraph(regExObject, this.paragraphs[location].Text, checkValue);
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
