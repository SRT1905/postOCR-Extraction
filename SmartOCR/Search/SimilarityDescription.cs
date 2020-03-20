namespace SmartOCR.Search
{
    using System;
    using SmartOCR.Search.SimilarityAlgorithms;

    /// <summary>
    /// Used to contain similarity properties between two strings.
    /// </summary>
    public class SimilarityDescription // TODO: find approach to calling specific similarity algorithm
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
        /// Default string metric algorithm is used.
        /// </summary>
        /// <param name="foundString">String, found within Word document.</param>
        /// <param name="checkString">String to compare with one, found within Word document.</param>
        public SimilarityDescription(string foundString, string checkString)
        {
            ValidateInput(foundString, checkString);
            this.InitializeFields(foundString, checkString);
        }

        /// <summary>
        /// Sets algorithm that is used to calculate string similarity for all instances, initialized after setting this property.
        /// </summary>
        public static ISimilarityAlgorithm SimilarityAlgorithm { private get; set; } = SimilarityAlgorithmSelector.GetAlgorithm("LevensteinAlgorithm");

        /// <summary>
        /// Gets percentage of closeness between two strings.
        /// </summary>
        public double Ratio { get; private set; }

        /// <summary>
        /// Gets source value.
        /// </summary>
        public string Source { get; private set; }

        /// <summary>
        /// Gets saved value of found string.
        /// </summary>
        public string Value { get; private set; }

        /// <summary>
        /// Checks whether calculated similarity ratio is sufficient enough to call two strings similar.
        /// </summary>
        /// <returns>true/false.</returns>
        public bool AreStringsSimilar()
        {
            return this.Ratio >= SimilarityRatioThreshold;
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"Value '{this.Value}' matches source '{this.Source}' by {this.Ratio * 100}%";
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

        private void InitializeFields(string foundString, string checkString)
        {
            if (foundString.Length - StringLengthOffset > checkString.Length ||
                foundString.Length + StringLengthOffset < checkString.Length)
            {
                return;
            }

            this.Ratio = SimilarityAlgorithm.GetStringSimilarity(foundString, checkString);
            this.Source = checkString;
            this.Value = foundString;
        }
    }
}