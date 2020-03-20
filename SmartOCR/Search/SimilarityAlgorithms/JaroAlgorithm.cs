namespace SmartOCR.Search.SimilarityAlgorithms
{
    using System;
    using System.Linq;

    /// <summary>
    /// Contains methods to calculate string similarity measure using Jaro similarity.
    /// </summary>
    public class JaroAlgorithm : ISimilarityAlgorithm
    {
        /// <inheritdoc/>
        public virtual double GetStringSimilarity(string leftString, string rightString)
        {
            return GetJaroSimilarity(leftString, rightString);
        }

        /// <summary>
        /// Calculates similarity ratio between strings using Jaro similarity measure.
        /// </summary>
        /// <param name="leftString">First compared string.</param>
        /// <param name="rightString">Second compared string.</param>
        /// <returns>Similarity ratio between two strings.</returns>
        protected static double GetJaroSimilarity(string leftString, string rightString)
        {
            return leftString != rightString
                ? leftString.Length != 0 && rightString.Length != 0
                    ? GetSimilarityFromValidStrings(leftString, rightString)
                    : 0.0
                : 1.0;
        }

        private static double GetSimilarityFromValidStrings(string leftString, string rightString)
        {
            // Hash for matches
            var hashLeft = new int[leftString.Length];
            var hashRight = new int[rightString.Length];
            var matchCount = GetMatchCount(leftString, rightString, hashLeft, hashRight);
            return matchCount == 0
                ? 0.0
                : CalculateSimilarityRatio(
                    leftString.Length,
                    rightString.Length,
                    matchCount,
                    GetCountOfTranspositions(leftString, rightString, hashLeft, hashRight));
        }

        private static double CalculateSimilarityRatio(int leftStringLength, int rightStringLength, int matchCount, double transpositions)
        {
            return (((double)matchCount / leftStringLength)
                + ((double)matchCount / rightStringLength)
                + ((matchCount - transpositions) / matchCount)) / 3.0;
        }

        private static int GetMatchCount(string leftString, string rightString, int[] hashLeft, int[] hashRight)
        {
            return leftString.Select((t, charIndex) => GetCountOfMatchesWithSingleChar(leftString, rightString, hashLeft, hashRight, charIndex)).Sum();
        }

        private static int GetCountOfMatchesWithSingleChar(string leftString, string rightString, int[] hashLeft, int[] hashRight, int charIndex)
        {
            Tuple<int, int> indexes = GetIndexesForMatchSearch(
                charIndex,
                GetMaxDistance(leftString, rightString),
                rightString.Length);
            for (int j = indexes.Item1; j < indexes.Item2; j++)
            {
                if (leftString[charIndex] == rightString[j] && hashRight[j] == 0)
                {
                    return ReturnSuccessfulMatch(hashLeft, hashRight, charIndex, j);
                }
            }

            return 0;
        }

        private static int ReturnSuccessfulMatch(int[] hashLeft, int[] hashRight, int leftCharIndex, int rightCharIndex)
        {
            hashLeft[leftCharIndex] = 1;
            hashRight[rightCharIndex] = 1;
            return 1;
        }

        private static Tuple<int, int> GetIndexesForMatchSearch(int leftStringCharIndex, int maxDistance, int rightStringLength)
        {
            return new Tuple<int, int>(
                Math.Max(0, leftStringCharIndex - maxDistance),
                Math.Min(rightStringLength, leftStringCharIndex + maxDistance + 1));
        }

        private static int GetMaxDistance(string leftString, string rightString)
        {
            // Maximum distance up to which matching is allowed
            return (int)Math.Floor((double)Math.Max(leftString.Length, rightString.Length) / 2) - 1;
        }

        private static double GetCountOfTranspositions(string leftString, string rightString, int[] hashLeft, int[] hashRight)
        {
            // Number of transpositions
            double transpositions = 0;
            int point = 0;

            // Count number of cases when two characters match but there is a third matched char between the indexes
            for (int i = 0; i < leftString.Length; i++)
            {
                if (hashLeft[i] == 1)
                {
                    point = FindMatchingPointInSecondString(hashRight, point);

                    if (leftString[i] != rightString[point++])
                    {
                        transpositions++;
                    }
                }
            }

            return transpositions / 2;
        }

        private static int FindMatchingPointInSecondString(int[] hashRight, int point)
        {
            // Find next matches char in second string
            while (hashRight[point] == 0)
            {
                point++;
            }

            return point;
        }
    }
}
