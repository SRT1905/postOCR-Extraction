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
            { 'b', '1' }, { 'f', '1' }, { 'p', '1' }, { 'v', '1' },
            { 'c', '2' }, { 'g', '2' }, { 'j', '2' }, { 'k', '2' },
            { 'q', '2' }, { 's', '2' }, { 'x', '2' }, { 'z', '2' },
            { 'd', '3' }, { 't', '3' },
            { 'l', '4' },
            { 'm', '5' }, { 'n', '5' },
            { 'r', '6' },
        };

        private static readonly HashSet<char> Vowels = new HashSet<char>()
        {
            'a', 'e', 'i', 'o', 'u', 'y',
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
                              .Where(singleChar => !(singleChar == 'h' || singleChar == 'w'))
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
