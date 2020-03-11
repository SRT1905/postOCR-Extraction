namespace SmartOCR
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Used to encode strings respresentation so that they can be matched despite minor differences in spelling.
    /// </summary>
    public struct Soundex
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
        /// Initializes a new instance of the <see cref="Soundex"/> struct.
        /// </summary>
        /// <param name="sourceValue">Original decoded string value.</param>
        public Soundex(string sourceValue)
        {
            this.SourceValue = sourceValue;
            this.EncodedValue = EncodeValue(sourceValue);
        }

        /// <summary>
        /// Gets original decoded string value.
        /// </summary>
        public string SourceValue { get; private set; }

        /// <summary>
        /// Gets encoded string value.
        /// </summary>
        public string EncodedValue { get; private set; }

        /// <summary>
        /// Returns encoded Soundex value of source string.
        /// </summary>
        /// <param name="value">Original decoded string.</param>
        /// <returns>And encoded Soundex value.</returns>
        public static string EncodeValue(string value)
        {
            string[] splittedValue = value.Split(' ');

            for (int i = 0; i < splittedValue.Length; i++)
            {
                splittedValue[i] = EncodeSingleWord(splittedValue[i].ToLower());
            }

            return string.Join(" ", splittedValue);
        }

        private static string EncodeSingleWord(string word)
        {
            List<char> sourceChars = word.ToList();
            char firstLetter = sourceChars[0];
            sourceChars = PrepareSourceChars(sourceChars);
            return Finalize(char.ToUpper(firstLetter), sourceChars);
        }

        private static List<char> PrepareSourceChars(List<char> sourceChars)
        {
            sourceChars = RemoveInvalidLetters(sourceChars);

            // sourceChars.Insert(0, firstLetter);
            sourceChars = DoIndexProcedures(sourceChars);
            sourceChars = RemoveVowels(sourceChars);
            return sourceChars;
        }

        private static string Finalize(char firstLetter, List<char> sourceChars)
        {
            sourceChars.Insert(0, firstLetter);
            while (sourceChars.Count < 4)
            {
                sourceChars.Add('0');
            }

            return new string(sourceChars.Take(4).ToArray());
        }

        private static List<char> RemoveInvalidLetters(List<char> sourceChars)
        {
            sourceChars = sourceChars.Skip(1).ToList();
            sourceChars.RemoveAll(item => item == 'h' || item == 'w');
            return sourceChars;
        }

        private static List<char> ReplaceCharsByIndexes(List<char> sourceChars)
        {
            for (int i = 0; i < sourceChars.Count; i++)
            {
                if (CharIndexes.ContainsKey(sourceChars[i]))
                {
                    sourceChars[i] = CharIndexes[sourceChars[i]];
                }
            }

            return sourceChars;
        }

        private static List<char> RemoveVowels(List<char> sourceChars)
        {
            sourceChars = sourceChars.Skip(1).ToList();
            sourceChars.RemoveAll(item => Vowels.Contains(item));
            return sourceChars;
        }

        private static List<char> TrimRepeatingIndexes(List<char> sourceChars)
        {
            return sourceChars.Distinct().ToList();
        }

        private static List<char> DoIndexProcedures(List<char> sourceChars)
        {
            sourceChars = ReplaceCharsByIndexes(sourceChars);
            return TrimRepeatingIndexes(sourceChars);
        }
    }
}
