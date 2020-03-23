namespace SmartOCR.Soundex
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Encodes string using default Soundex encoding algorithm.
    /// </summary>
    public class DefaultSoundexEncoder : Soundex
    {
        private static readonly Dictionary<char, char> CharIndexes = new Dictionary<char, char>()
        {
            { 'B', '1' }, { 'F', '1' }, { 'P', '1' }, { 'V', '1' },
            { 'C', '2' }, { 'G', '2' }, { 'J', '2' }, { 'K', '2' },
            { 'Q', '2' }, { 'S', '2' }, { 'X', '2' }, { 'Z', '2' },
            { 'D', '3' }, { 'T', '3' },
            { 'L', '4' },
            { 'M', '5' }, { 'N', '5' },
            { 'R', '6' },
        };

        private static readonly HashSet<char> Vowels = new HashSet<char>()
        {
            'A', 'E', 'I', 'O', 'U', 'Y',
        };

        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultSoundexEncoder"/> class.
        /// </summary>
        /// <param name="value">A string to encode.</param>
        public DefaultSoundexEncoder(string value)
        {
            this.SourceValue = value;
            this.EncodedValue = this.EncodeValue(value);
        }

        /// <inheritdoc/>
        protected override string EncodeSingleWord(string word)
        {
            List<char> sourceChars = word.ToList();
            char firstLetter = sourceChars[0];
            sourceChars = PrepareSourceChars(sourceChars);
            return Finalize(char.ToUpper(firstLetter), sourceChars);
        }

        private static void AddZerosToRight(List<char> sourceChars)
        {
            while (sourceChars.Count < 4)
            {
                sourceChars.Add('0');
            }
        }

        private static void TryReplaceSingleChar(List<char> sourceChars, int index)
        {
            sourceChars[index] = CharIndexes.ContainsKey(sourceChars[index])
                                ? CharIndexes[sourceChars[index]]
                                : sourceChars[index];
        }

        private static List<char> PrepareSourceChars(List<char> sourceChars)
        {
            sourceChars = RemoveInvalidLetters(sourceChars);
            sourceChars = DoIndexProcedures(sourceChars);
            sourceChars = RemoveVowels(sourceChars);
            return sourceChars;
        }

        private static List<char> RemoveInvalidLetters(List<char> sourceChars)
        {
            return sourceChars.Skip(1)
                              .Where(singleChar => !(singleChar == 'H' || singleChar == 'W'))
                              .ToList();
        }

        private static List<char> ReplaceCharsByIndexes(List<char> sourceChars)
        {
            for (int i = 0; i < sourceChars.Count; i++)
            {
                TryReplaceSingleChar(sourceChars, i);
            }

            return sourceChars;
        }

        private static List<char> RemoveVowels(List<char> sourceChars)
        {
            sourceChars.RemoveAll(item => Vowels.Contains(item));
            return sourceChars;
        }

        private static List<char> DoIndexProcedures(List<char> sourceChars)
        {
            sourceChars = ReplaceCharsByIndexes(sourceChars);
            return TrimRepeatingIndexes(sourceChars);
        }

        private static string Finalize(char firstLetter, List<char> sourceChars)
        {
            sourceChars.Insert(0, firstLetter);
            AddZerosToRight(sourceChars);

            return new string(sourceChars.Take(4).ToArray());
        }
    }
}
