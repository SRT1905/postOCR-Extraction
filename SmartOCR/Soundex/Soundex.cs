namespace SmartOCR.Soundex
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to encode strings representation so that they can be matched despite minor differences in spelling.
    /// </summary>
    public abstract class Soundex
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Soundex"/> class.
        /// </summary>
        /// <param name="value">A string to encode.</param>
        protected Soundex(string value)
        {
            this.SourceValue = value;
            this.EncodedValue = this.EncodeValue(value);
        }

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
            var splitValue = Regex.Split(value, "([^а-яёa-z]+)", RegexOptions.IgnoreCase);
            for (var index = 0; index < splitValue.Length; index++)
            {
                if (Regex.IsMatch(splitValue[index], "[а-яёa-z]", RegexOptions.IgnoreCase))
                {
                    splitValue[index] = this.EncodeSingleWord(splitValue[index].ToUpper());
                }
            }

            return string.Join(string.Empty, splitValue);
        }

        /// <summary>
        /// Removes consecutive characters from List.
        /// </summary>
        /// <param name="sourceChars">A source char list.</param>
        /// <returns>A list of chars with trimmed consecutive letters.</returns>
        protected static List<char> TrimRepeatingIndexes(List<char> sourceChars)
        {
            var sb = new StringBuilder();
            new string(sourceChars.ToArray()).Aggregate(char.MinValue, (current, item) => TryAddCharToBuilder(sb, item, current));
            return sb.ToString().ToList();
        }

        /// <summary>
        /// Performs encoding of single word, using specific encoding algorithm.
        /// </summary>
        /// <param name="word">A word to encode.</param>
        /// <returns>An encoded string.</returns>
        protected abstract string EncodeSingleWord(string word);

        private static char TryAddCharToBuilder(StringBuilder sb, char charToAdd, char previousChar)
        {
            if (charToAdd == previousChar)
            {
                return previousChar;
            }

            sb.Append(charToAdd);
            previousChar = charToAdd;
            return previousChar;
        }
    }
}
