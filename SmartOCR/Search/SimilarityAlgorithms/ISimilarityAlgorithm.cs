namespace SmartOCR.Search.SimilarityAlgorithms
{
    /// <summary>
    /// Defines an interface for all classes that calculate similarity measures between two strings.
    /// </summary>
    public interface ISimilarityAlgorithm
    {
        /// <summary>
        /// Returns similarity measure between two strings, calculated by algorithm.
        /// </summary>
        /// <param name="leftString">First compared string.</param>
        /// <param name="rightString">Second compared string.</param>
        /// <returns>Similarity ratio between two strings.</returns>
        double GetStringSimilarity(string leftString, string rightString);
    }
}
