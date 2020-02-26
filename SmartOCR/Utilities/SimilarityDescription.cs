namespace SmartOCR
{
    using System;
    using System.Globalization;

    /// <summary>
    /// Used to contain similarity properties between two strings.
    /// </summary>
    public class SimilarityDescription
    {
        /// <summary>
        /// Lower bound of valid similarity ratio.
        /// </summary>
        private const double SimilarityRatioThreshold = 0.66;

        /// <summary>
        /// Margin, by which found string length may vary in comparison with check string.
        /// </summary>
        private const byte StringLengthOffset = 1;

        /// <summary>
        /// Initializes a new instance of the <see cref="SimilarityDescription"/> class.
        /// Instance is initialized with strings to be compared.
        /// </summary>
        /// <param name="foundString">String, found within Word document.</param>
        /// <param name="checkString">String to compare with one, found within Word document.</param>
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

            if (foundString.Length - StringLengthOffset <= checkString.Length &&
                foundString.Length + StringLengthOffset >= checkString.Length)
            {
                this.Ratio = this.GetStringSimilarity(foundString, checkString);
                this.Value = foundString;
            }
        }

        /// <summary>
        /// Gets percentage of closeness between two strings.
        /// </summary>
        public double Ratio { get; }

        /// <summary>
        /// Gets saved value of found string.
        /// </summary>
        public string Value { get; }

        /// <summary>
        /// Checks whether calculated similiarity ratio is sufficient enough to call two strings similar.
        /// </summary>
        /// <returns>true/false.</returns>
        public bool CheckStringSimilarity()
        {
            return this.Ratio >= SimilarityRatioThreshold;
        }

        /// <summary>
        /// Calculates number of string operations, which would take to transform one string to another.
        /// </summary>
        /// <param name="longString">String with bigger length, comparing with other.</param>
        /// <param name="shortString">String with smaller length, comparing with other.</param>
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

            for (int i = 0; i <= longLength; costs[i][0] = i++)
            {
            }

            for (int j = 0; j <= shortLength; costs[0][j] = j++)
            {
            }

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

        /// <summary>
        /// Prepares strings for similarity testing and calculates ratio.
        /// </summary>
        /// <param name="leftString">First string to check.</param>
        /// <param name="rightString">Second string to check.</param>
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
    }
}