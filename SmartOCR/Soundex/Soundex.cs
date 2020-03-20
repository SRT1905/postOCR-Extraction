namespace SmartOCR.Soundex
{
    /// <summary>
    /// Used to encode strings representation so that they can be matched despite minor differences in spelling.
    /// </summary>
    public abstract class Soundex
    {
        /// <summary>
        /// Gets or sets original decoded string value.
        /// </summary>
        public string SourceValue { get; protected set; }

        /// <summary>
        /// Gets or sets encoded string value.
        /// </summary>
        public string EncodedValue { get; protected set; }

        /// <summary>
        /// Returns encoded Soundex value of source string.
        /// </summary>
        /// <param name="value">Original decoded string.</param>
        /// <returns>And encoded Soundex value.</returns>
        public string EncodeValue(string value)
        {
            string[] splittedValue = value.Split(' ');

            for (int i = 0; i < splittedValue.Length; i++)
            {
                splittedValue[i] = this.EncodeSingleWord(splittedValue[i].ToLower());
            }

            return string.Join(" ", splittedValue);
        }

        /// <summary>
        /// Performs encoding of single word, using specific encoding algorithm.
        /// </summary>
        /// <param name="word">A word to encode.</param>
        /// <returns>An encoded string.</returns>
        protected abstract string EncodeSingleWord(string word);
    }
}
