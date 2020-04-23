namespace SmartOCR.Search.SimilarityAlgorithms
{
    using System;

    /// <summary>
    /// Contains methods to calculate string similarity measure using Jaro-Winkler similarity.
    /// </summary>
    public class JaroWinklerAlgorithm : JaroAlgorithm
    {
        private const double JaroSimilarityThreshold = 0.7;

        /// <inheritdoc/>
        public override double GetStringSimilarity(string leftString, string rightString)
        {
            return GetJaroWinklerDistance(leftString, rightString);
        }

        private static double GetJaroWinklerDistance(string leftString, string rightString)
        {
            double jaroDistance = GetJaroSimilarity(leftString, rightString);

            return jaroDistance > JaroSimilarityThreshold
                ? CalculateJaroWinklerMeasure(leftString, rightString, jaroDistance)
                : jaroDistance;
        }

        private static double CalculateJaroWinklerMeasure(string leftString, string rightString, double jaroDistance)
        {
            return jaroDistance + (0.1 * GetLengthOfCommonPrefix(leftString, rightString) * (1 - jaroDistance));
        }

        private static int GetLengthOfCommonPrefix(string leftString, string rightString)
        {
            int prefix = 0;
            for (int i = 0; i < Math.Min(leftString.Length, rightString.Length); i++)
            {
                if (leftString[i] == rightString[i])
                {
                    prefix++;
                }
                else
                {
                    break;
                }
            }

            return Math.Min(4, prefix);
        }
    }
}