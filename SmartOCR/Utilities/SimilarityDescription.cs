using System;
using System.Globalization;

namespace SmartOCR
{
    /// <summary>
    /// Used to contain similarity properties between two strings.
    /// </summary>
    public class SimilarityDescription
    {
        #region Prviate constants
        /// <summary>
        /// Lower bound of valid similarity ratio.
        /// </summary>
        private const double similarityRatioThreshold = 0.66;
        /// <summary>
        /// Margin, by which found string length may vary in comparison with check string.
        /// </summary>
        private const byte stringLengthOffset = 1;
        #endregion

        #region Properties
        /// <summary>
        /// Represents percentage of closeness between two strings.
        /// </summary>
        public double Ratio { get; }
        /// <summary>
        /// Saved value of found string.
        /// </summary>
        public string Value { get; }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new <see cref="SimilarityDescription"/> instance with strings to be compared.
        /// </summary>
        /// <param name="foundString">String, found within Word document.</param>
        /// <param name="checkString"></param>
        public SimilarityDescription(string foundString, string checkString)
        {
            if (foundString == null)
            {
                throw new ArgumentNullException(nameof(foundString));
            }
            if (checkString == null)
            {
                throw new ArgumentNullException(nameof(checkString));
            }
            if (foundString.Length - stringLengthOffset <= checkString.Length &&
                foundString.Length + stringLengthOffset >= checkString.Length)
            {
                Ratio = GetStringSimilarity(foundString, checkString);
                Value = foundString;
            }
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Checks whether calculated similiarity ratio is sufficient enough to call two strings similar.
        /// </summary>
        /// <returns>true/false</returns>
        public bool CheckStringSimilarity()
        {
            return Ratio >= similarityRatioThreshold;
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Prepares strings for similarity testing and calculates ratio.
        /// </summary>
        /// <param name="leftString"></param>
        /// <param name="rightString"></param>
        /// <returns>Similarity ratio.</returns>
        private double GetStringSimilarity(string leftString, string rightString)
        {
            string shortString;
            string longString;
            if (leftString.Length < rightString.Length)
            {
                shortString = leftString;
                longString = rightString;
            }
            else
            {
                shortString = rightString;
                longString = leftString;
            }
            return longString.ToLower(CultureInfo.CurrentCulture) == shortString.ToLower(CultureInfo.CurrentCulture)
                ? 1
                : (longString.Length - ComputeLevensteinDistance(longString, shortString)) / longString.Length;
        }
        #endregion

        #region Private static methods
        /// <summary>
        /// Calculates number of string operations, which would take to transform one string to another.
        /// </summary>
        /// <param name="longString"></param>
        /// <param name="shortString"></param>
        /// <returns>Number of required operations.</returns>
        private static double ComputeLevensteinDistance(string longString, string shortString)
        {
            int longLength = longString.Length;
            int shortLength = shortString.Length;

            int[][] costs = new int[longLength + 1][];
            for (int i = 0; i < costs.Length; i++)
            {
                costs[i] = new int[shortLength + 1];
            }
            if (longLength == 0)
            {
                return shortLength;
            }

            if (shortLength == 0)
            {
                return longLength;
            }

            for (int i = 0; i <= longLength; costs[i][0] = i++) ;
            for (int j = 0; j <= shortLength; costs[0][j] = j++) ;
            for (int i = 1; i <= longLength; i++)
            {
                for (int j = 1; j <= shortLength; j++)
                {
                    int cost = (shortString[j - 1] == longString[i - 1]) ? 0 : 1;
                    costs[i][j] = Math.Min(
                        Math.Min(
                            costs[i - 1][j] + 1,
                            costs[i][j - 1] + 1),
                        costs[i - 1][j - 1] + cost);
                }
            }
            return costs[longLength][shortLength];
        }
        #endregion
    }
}