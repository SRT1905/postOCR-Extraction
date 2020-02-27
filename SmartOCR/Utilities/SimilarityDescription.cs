namespace SmartOCR
{
    using System;

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
            ValidateInput(foundString, checkString);
            this.InitializeFields(foundString, checkString);
        }

        /// <summary>
        /// Gets percentage of closeness between two strings.
        /// </summary>
        public double Ratio { get; private set; }

        /// <summary>
        /// Gets saved value of found string.
        /// </summary>
        public string Value { get; private set; }

        /// <summary>
        /// Checks whether calculated similiarity ratio is sufficient enough to call two strings similar.
        /// </summary>
        /// <returns>true/false.</returns>
        public bool CheckStringSimilarity()
        {
            return this.Ratio >= SimilarityRatioThreshold;
        }

        private static void ValidateInput(string foundString, string checkString)
        {
            if (foundString == null)
            {
                throw new ArgumentNullException(nameof(foundString));
            }

            if (checkString == null)
            {
                throw new ArgumentNullException(nameof(checkString));
            }
        }

        /// <summary>
        /// Calculates number of string operations, which would take to transform one string to another.
        /// </summary>
        /// <param name="longString">String with bigger length, comparing with other.</param>
        /// <param name="shortString">String with smaller length, comparing with other.</param>
        /// <returns>Number of required operations.</returns>
        private static double ComputeLevensteinDistance(string longString, string shortString)
        {
            if (longString.Length == 0 || shortString.Length == 0)
            {
                return Math.Max(shortString.Length, longString.Length);
            }

            return InitializeCostsArray(longString, shortString)[longString.Length][shortString.Length];
        }

        private static void RecalculateCosts(string longString, string shortString, int[][] costs)
        {
            for (int i = 1; i <= longString.Length; i++)
            {
                for (int j = 1; j <= shortString.Length; j++)
                {
                    RecalculateSingleCost(longString, shortString, costs, i, j);
                }
            }
        }

        private static void RecalculateSingleCost(string longString, string shortString, int[][] costs, int i, int j)
        {
            int cost = Convert.ToInt32(shortString[j - 1] != longString[i - 1]);
            costs[i][j] = Math.Min(
                Math.Min(costs[i - 1][j] + 1, costs[i][j - 1] + 1),
                costs[i - 1][j - 1] + cost);
        }

        private static int[][] InitializeCostsArray(string longString, string shortString)
        {
            int[][] costs = new int[longString.Length + 1][];
            for (int i = 0; i < costs.Length; i++)
            {
                costs[i] = new int[shortString.Length + 1];
            }

            FillArrayWithUnits(longString, shortString, costs);
            RecalculateCosts(longString, shortString, costs);
            return costs;
        }

        private static void FillArrayWithUnits(string longString, string shortString, int[][] costs)
        {
            for (int i = 0; i <= longString.Length; i++)
            {
                costs[i][0] = i + 1;
            }

            for (int j = 0; j <= shortString.Length; j++)
            {
                costs[0][j] = j + 1;
            }
        }

        private static double CalculateSimilarityRatio(string shortString, string longString)
        {
            return (longString.Length - ComputeLevensteinDistance(longString, shortString)) / longString.Length;
        }

        private void InitializeFields(string foundString, string checkString)
        {
            if (foundString.Length - StringLengthOffset <= checkString.Length &&
                foundString.Length + StringLengthOffset >= checkString.Length)
            {
                this.Ratio = this.GetStringSimilarity(foundString, checkString);
                this.Value = foundString;
            }
        }

        /// <summary>
        /// Prepares strings for similarity testing and calculates ratio.
        /// </summary>
        /// <param name="leftString">First string to check.</param>
        /// <param name="rightString">Second string to check.</param>
        /// <returns>Similarity ratio.</returns>
        private double GetStringSimilarity(string leftString, string rightString)
        {
            string shortString = leftString.ToLower();
            string longString = rightString.ToLower();
            if (leftString.Length > rightString.Length)
            {
                shortString = rightString.ToLower();
                longString = leftString.ToLower();
            }

            return longString == shortString
                ? 1
                : CalculateSimilarityRatio(shortString, longString);
        }
    }
}