namespace SmartOCR
{
    using System;

    /// <summary>
    /// Contains methods to calculate string similarity measure using Levenstein distance.
    /// </summary>
    public class LevensteinAlgorithm : ISimilarityAlgorithm
    {
        /// <inheritdoc/>
        public double GetStringSimilarity(string leftString, string rightString)
        {
            Tuple<string, string> stringPair = DistinguishStringByLength(leftString, rightString);

            return stringPair.Item1 == stringPair.Item2
                ? 1
                : CalculateSimilarityRatio(stringPair.Item1, stringPair.Item2);
        }

        private static Tuple<string, string> DistinguishStringByLength(string leftString, string rightString)
        {
            string shortString = leftString.Length > rightString.Length
                ? rightString.ToLower()
                : leftString.ToLower();

            string longString = leftString.Length > rightString.Length
                ? leftString.ToLower()
                : rightString.ToLower();
            return new Tuple<string, string>(shortString, longString);
        }

        private static double CalculateSimilarityRatio(string shortString, string longString)
        {
            return (longString.Length - ComputeLevensteinDistance(shortString, longString)) / longString.Length;
        }

        /// <summary>
        /// Calculates number of string operations, which would take to transform one string to another.
        /// </summary>
        /// <param name="shortString">String with smaller length, comparing with other.</param>
        /// <param name="longString">String with bigger length, comparing with other.</param>
        /// <returns>Number of required operations.</returns>
        private static double ComputeLevensteinDistance(string shortString, string longString)
        {
            return longString.Length == 0 || shortString.Length == 0
                ? Math.Max(shortString.Length, longString.Length)
                : InitializeCostsArray(shortString, longString)[longString.Length][shortString.Length];
        }

        private static int[][] InitializeCostsArray(string shortString, string longString)
        {
            var costs = new int[longString.Length + 1][];
            for (int i = 0; i < costs.Length; i++)
            {
                costs[i] = new int[shortString.Length + 1];
            }

            FillArrayWithUnits(shortString, longString, costs);
            RecalculateCosts(shortString, longString, costs);
            return costs;
        }

        private static void FillArrayWithUnits(string shortString, string longString, int[][] costs)
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

        private static void RecalculateCosts(string shortString, string longString, int[][] costs)
        {
            for (int i = 1; i <= longString.Length; i++)
            {
                for (int j = 1; j <= shortString.Length; j++)
                {
                    RecalculateSingleCost(shortString, longString, costs, i, j);
                }
            }
        }

        private static void RecalculateSingleCost(string shortString, string longString, int[][] costs, int i, int j)
        {
            int cost = Convert.ToInt32(shortString[j - 1] != longString[i - 1]);
            costs[i][j] = Math.Min(
                Math.Min(costs[i - 1][j] + 1, costs[i][j - 1] + 1),
                costs[i - 1][j - 1] + cost);
        }
    }
}
