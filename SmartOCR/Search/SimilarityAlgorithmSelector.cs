namespace SmartOCR.Search
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using SmartOCR.Search.SimilarityAlgorithms;

    /// <summary>
    /// Used to access similarity algorithms from current assembly.
    /// </summary>
    public static class SimilarityAlgorithmSelector
    {
        private static readonly List<Tuple<string, ISimilarityAlgorithm>> Algorithms = GetAlgorithms();

        /// <summary>
        /// Gets an instance, implementing <see cref="ISimilarityAlgorithm"/>, by its name.
        /// </summary>
        /// <param name="algorithmName">An algorithm title.</param>
        /// <returns>An instance, implementing <see cref="ISimilarityAlgorithm"/>.</returns>
        public static ISimilarityAlgorithm GetAlgorithm(string algorithmName) => Algorithms
            .Where(algorithmPair => algorithmPair.Item1 == algorithmName)
            .Select(algorithmPair => algorithmPair.Item2)
            .FirstOrDefault();

        private static List<Tuple<string, ISimilarityAlgorithm>> GetAlgorithms()
        {
            return typeof(ISimilarityAlgorithm).Assembly.GetTypes()
                .Where(singleType =>
                    typeof(ISimilarityAlgorithm).IsAssignableFrom(singleType) && !singleType.IsInterface)
                .Select(singleType =>
                    Tuple.Create(singleType.Name, (ISimilarityAlgorithm)Activator.CreateInstance(singleType))).ToList();
        }
    }
}