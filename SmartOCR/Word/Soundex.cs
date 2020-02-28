namespace SmartOCR
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Used to encode strings respresentation so that they can be matched despite minor differences in spelling.
    /// </summary>
    public class Soundex
    {
        private static readonly Dictionary<char, char> CharIndexes = new Dictionary<char, char>()
        {
            { 'b', '1' }, { 'f', '1' }, { 'p', '1' }, { 'v', '1' },
            { 'c', '2' }, { 'g', '2' }, { 'j', '2' }, { 'k', '2' }, { 'q', '2' }, { 's', '2' }, { 'x', '2' }, { 'z', '2' },
            { 'd', '3' }, { 't', '3' },
            { 'l', '4' },
            { 'm', '5' }, { 'n', '5' },
            { 'r', '6' },
        };

        private static readonly HashSet<char> Vowels = new HashSet<char>()
        {
            'a', 'e', 'i', 'o', 'u', 'y',
        };

        private char firstLetter;
        private List<char> sourceChars;

        /// <summary>
        /// Initializes a new instance of the <see cref="Soundex"/> class.
        /// </summary>
        /// <param name="sourceValue">Original decoded string value.</param>
        public Soundex(string sourceValue)
        {
            this.SourceValue = sourceValue;
            this.EncodedValue = this.EncodeSourceValue(sourceValue);
        }

        /// <summary>
        /// Gets original decoded string value.
        /// </summary>
        public string SourceValue { get; private set; }

        /// <summary>
        /// Gets encoded string value.
        /// </summary>
        public string EncodedValue { get; private set; }

        private string Finalize()
        {
            while (this.sourceChars.Count < 4)
            {
                this.sourceChars.Add('0');
            }

            return new string(this.sourceChars.Take(4).ToArray());
        }

        private void RemoveInvalidLetters()
        {
            this.sourceChars = this.sourceChars.Skip(1).ToList();
            this.sourceChars.RemoveAll(item => item == 'h' || item == 'w');
        }

        private void ReplaceCharsByIndexes()
        {
            for (int i = 0; i < this.sourceChars.Count; i++)
            {
                if (CharIndexes.ContainsKey(this.sourceChars[i]))
                {
                    this.sourceChars[i] = CharIndexes[this.sourceChars[i]];
                }
            }
        }

        private void RemoveVowels()
        {
            this.sourceChars = this.sourceChars.Skip(1).ToList();
            this.sourceChars.RemoveAll(item => Vowels.Contains(item));
        }

        private void TrimRepeatingIndexes()
        {
            int i = 1;
            while (i < this.sourceChars.Count)
            {
                if (this.sourceChars[i] == this.sourceChars[i - 1] && CharIndexes.ContainsValue(this.sourceChars[i]))
                {
                    this.sourceChars.RemoveAt(i);
                }
                else
                {
                    i++;
                }
            }
        }

        private string EncodeSourceValue(string sourceValue)
        {
            this.sourceChars = sourceValue.ToList();
            this.firstLetter = this.sourceChars[0];
            this.RemoveInvalidLetters();
            this.sourceChars.Insert(0, this.firstLetter);
            this.DoIndexProcedures();
            this.RemoveVowels();
            this.sourceChars.Insert(0, char.ToUpper(this.firstLetter));
            return this.Finalize();
        }

        private void DoIndexProcedures()
        {
            this.ReplaceCharsByIndexes();
            this.TrimRepeatingIndexes();
        }
    }
}
