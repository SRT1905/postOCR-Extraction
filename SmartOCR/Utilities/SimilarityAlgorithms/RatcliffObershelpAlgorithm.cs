namespace SmartOCR
{
    using System;
    using System.Linq;

    /// <summary>
    /// Represents class for calculating string similarity using Ratcliff-Obershelp algorithm.
    /// </summary>
    public class RatcliffObershelpAlgorithm : ISimilarityAlgorithm
    {
        /// <inheritdoc/>
        public double GetStringSimilarity(string leftString, string rightString)
        {
            return 2 * Convert.ToDouble(leftString.Intersect(rightString).Count()) /
                Convert.ToDouble(leftString.Length + rightString.Length);
        }
    }
}
