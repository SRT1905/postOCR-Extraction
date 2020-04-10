namespace SmartOCR.Soundex
{
    using System.Collections.Generic;
    using System.Linq;
    using SmartOCR.Utilities;

    /// <summary>
    /// Encodes string using Daitch-Mokotoff soundex algorithm.
    /// </summary>
    public class DaitchMokotoffSoundexEncoder : Soundex
    {
        /// <summary>
        /// Letter combination should be replaced by code if combination
        /// [0] - is a start of encoded value,
        /// [1] - is before a vowel,
        /// [2] - is in any other situation.
        /// Symbol '|' is used to separate possible codes for combination.
        /// </summary>
        private static readonly Dictionary<string, string[]> CodingChart =
            new Dictionary<string, string[]>
            {
                ["A"] = new[] { "0", "_", "_" },
                ["AI"] = new[] { "0", "1", "_" },
                ["AJ"] = new[] { "0", "1", "_" },
                ["AU"] = new[] { "0", "7", "_" },
                ["AY"] = new[] { "0", "1", "_" },

                ["B"] = new[] { "7", "7", "7" },

                ["C"] = new[] { "5|4", "5|4", "5|4" },
                ["CH"] = new[] { "5|4", "5|4", "5|4" },
                ["CHS"] = new[] { "5", "54", "54" },
                ["CK"] = new[] { "5|45", "5|45", "5|45" },
                ["CS"] = new[] { "4", "4", "4" },
                ["CSZ"] = new[] { "4", "4", "4" },
                ["CZ"] = new[] { "4", "4", "4" },
                ["CZS"] = new[] { "4", "4", "4" },

                ["D"] = new[] { "3", "3", "3" },
                ["DRS"] = new[] { "4", "4", "4" },
                ["DRZ"] = new[] { "4", "4", "4" },
                ["DS"] = new[] { "4", "4", "4" },
                ["DSH"] = new[] { "4", "4", "4" },
                ["DSZ"] = new[] { "4", "4", "4" },
                ["DT"] = new[] { "3", "3", "3" },
                ["DZ"] = new[] { "4", "4", "4" },
                ["DZH"] = new[] { "4", "4", "4" },
                ["DZS"] = new[] { "4", "4", "4" },

                ["E"] = new[] { "0", "_", "_" },
                ["EI"] = new[] { "0", "1", "_" },
                ["EJ"] = new[] { "0", "1", "_" },
                ["EU"] = new[] { "1", "1", "_" },
                ["EY"] = new[] { "0", "1", "_" },

                ["F"] = new[] { "7", "7", "7" },
                ["FB"] = new[] { "7", "7", "7" },

                ["G"] = new[] { "5", "5", "5" },

                ["H"] = new[] { "5", "5", "_" },

                ["I"] = new[] { "0", "_", "_" },
                ["IA"] = new[] { "1", "_", "_" },
                ["IE"] = new[] { "1", "_", "_" },
                ["IO"] = new[] { "1", "_", "_" },
                ["IU"] = new[] { "1", "_", "_" },

                ["J"] = new[] { "1|4", "_|4", "_|4" },

                ["K"] = new[] { "5", "5", "5" },
                ["KH"] = new[] { "5", "5", "5" },
                ["KS"] = new[] { "5", "54", "54" },

                ["L"] = new[] { "8", "8", "8" },

                ["M"] = new[] { "6", "6", "6" },
                ["MN"] = new[] { "6_6", "6_6", "6_6" },

                ["N"] = new[] { "6", "6", "6" },
                ["NM"] = new[] { "6_6", "6_6", "6_6" },

                ["O"] = new[] { "0", "_", "_" },
                ["OI"] = new[] { "0", "1", "_" },
                ["OJ"] = new[] { "0", "1", "_" },
                ["OY"] = new[] { "0", "1", "_" },

                ["P"] = new[] { "7", "7", "7" },
                ["PF"] = new[] { "7", "7", "7" },
                ["PH"] = new[] { "7", "7", "7" },

                ["Q"] = new[] { "5", "5", "5" },

                ["R"] = new[] { "9", "9", "9" },
                ["RS"] = new[] { "94|4", "94|4", "94|4" },
                ["RZ"] = new[] { "94|4", "94|4", "94|4" },

                ["S"] = new[] { "4", "4", "4" },
                ["SC"] = new[] { "2", "4", "4" },
                ["SCH"] = new[] { "4", "4", "4" },
                ["SCHD"] = new[] { "2", "43", "43" },
                ["SCHT"] = new[] { "2", "43", "43" },
                ["SCHTCH"] = new[] { "2", "4", "4" },
                ["SCHTSCH"] = new[] { "2", "4", "4" },
                ["SCHTSH"] = new[] { "2", "4", "4" },
                ["SD"] = new[] { "2", "43", "43" },
                ["SH"] = new[] { "4", "4", "4" },
                ["SHCH"] = new[] { "2", "4", "4" },
                ["SHD"] = new[] { "2", "43", "43" },
                ["SHT"] = new[] { "2", "43", "43" },
                ["SHTCH"] = new[] { "2", "4", "4" },
                ["SHTSH"] = new[] { "2", "4", "4" },
                ["ST"] = new[] { "2", "43", "43" },
                ["STCH"] = new[] { "2", "4", "4" },
                ["STRS"] = new[] { "2", "4", "4" },
                ["STRZ"] = new[] { "2", "4", "4" },
                ["STSCH"] = new[] { "2", "4", "4" },
                ["STSH"] = new[] { "2", "4", "4" },
                ["SZ"] = new[] { "4", "4", "4" },
                ["SZCS"] = new[] { "2", "4", "4" },
                ["SZCZ"] = new[] { "2", "4", "4" },
                ["SZD"] = new[] { "2", "43", "43" },
                ["SZT"] = new[] { "2", "43", "43" },

                ["T"] = new[] { "3", "3", "3" },
                ["TC"] = new[] { "4", "4", "4" },
                ["TCH"] = new[] { "4", "4", "4" },
                ["TH"] = new[] { "3", "3", "3" },
                ["THS"] = new[] { "4", "4", "4" },
                ["TRS"] = new[] { "4", "4", "4" },
                ["TRZ"] = new[] { "4", "4", "4" },
                ["TS"] = new[] { "4", "4", "4" },
                ["TSCH"] = new[] { "4", "4", "4" },
                ["TSH"] = new[] { "4", "4", "4" },
                ["TSZ"] = new[] { "4", "4", "4" },
                ["TTCH"] = new[] { "4", "4", "4" },
                ["TTS"] = new[] { "4", "4", "4" },
                ["TTSCH"] = new[] { "4", "4", "4" },
                ["TTSZ"] = new[] { "4", "4", "4" },
                ["TTZ"] = new[] { "4", "4", "4" },
                ["TZ"] = new[] { "4", "4", "4" },
                ["TZS"] = new[] { "4", "4", "4" },

                ["U"] = new[] { "0", "_", "_" },
                ["UE"] = new[] { "0", "_", "_" },
                ["UI"] = new[] { "0", "1", "_" },
                ["UJ"] = new[] { "0", "1", "_" },
                ["UY"] = new[] { "0", "1", "_" },

                ["V"] = new[] { "7", "7", "7" },

                ["W"] = new[] { "7", "7", "7" },

                ["X"] = new[] { "5", "54", "54" },

                ["Y"] = new[] { "1", "_", "_" },

                ["Z"] = new[] { "4", "4", "4" },
                ["ZD"] = new[] { "2", "43", "43" },
                ["ZDZ"] = new[] { "2", "4", "4" },
                ["ZDZH"] = new[] { "2", "4", "4" },
                ["ZH"] = new[] { "4", "4", "4" },
                ["ZHD"] = new[] { "2", "43", "43" },
                ["ZHDZH"] = new[] { "2", "4", "4" },
                ["ZS"] = new[] { "4", "4", "4" },
                ["ZSCH"] = new[] { "4", "4", "4" },
                ["ZSH"] = new[] { "4", "4", "4" },
            };

        private static readonly Dictionary<char, string[]> CombinationsByFirstLetter =
            new Dictionary<char, string[]>
        {
            ['A'] = new[] { "AI", "AJ", "AU", "AY", "A" },
            ['B'] = new[] { "B" },
            ['C'] = new[]
            {
                "CHS", "CSZ", "CZS", "CH",
                "CK", "CS", "CZ", "C",
            },
            ['D'] = new[]
            {
                "DRS", "DRZ", "DSH", "DSZ",
                "DZH", "DZS", "DS", "DT", "DZ", "D",
            },
            ['E'] = new[] { "EI", "EJ", "EU", "EY", "E" },
            ['F'] = new[] { "FB", "F" },
            ['G'] = new[] { "G", },
            ['H'] = new[] { "H", },
            ['I'] = new[] { "IA", "IE", "IO", "IU", "I" },
            ['J'] = new[] { "J", },
            ['K'] = new[] { "KH", "KS", "K" },
            ['L'] = new[] { "L", },
            ['M'] = new[] { "MN", "M" },
            ['N'] = new[] { "NM", "N" },
            ['O'] = new[] { "OI", "OJ", "OY", "O" },
            ['P'] = new[] { "PF", "PH", "P" },
            ['Q'] = new[] { "Q", },
            ['R'] = new[] { "RS", "RZ", "R" },
            ['S'] = new[]
            {
                "SCHTSCH", "SCHTCH", "SCHTSH", "SHTCH", "SHTSH",
                "STSCH", "SCHD", "SCHT", "SHCH", "STCH", "STRS",
                "STRZ", "STSH", "SZCS", "SZCZ", "SCH", "SHD", "SHT",
                "SZD", "SZT", "SC", "SD", "SH", "ST", "SZ", "S",
            },
            ['T'] = new[]
            {
                "TTSCH", "TSCH", "TTCH", "TTSZ", "TCH",
                "THS", "TRS", "TRZ", "TSH", "TSZ",
                "TTS", "TTZ", "TZS", "TC", "TH",
                "TS", "TZ", "T",
            },
            ['U'] = new[] { "UE", "UI", "UJ", "UY", "U" },
            ['V'] = new[] { "V", },
            ['W'] = new[] { "W", },
            ['X'] = new[] { "X", },
            ['Y'] = new[] { "Y", },
            ['Z'] = new[]
            {
                "ZHDZH", "ZDZH", "ZSCH", "ZDZ", "ZHD",
                "ZSH", "ZD", "ZH", "ZS", "Z",
            },
        };

        private static readonly HashSet<char> CharsToRemove = new HashSet<char>
        {
            'A', 'E', 'I', 'J', 'O', 'U', 'Y',
        };

        /// <summary>
        /// Initializes a new instance of the <see cref="DaitchMokotoffSoundexEncoder"/> class.
        /// </summary>
        /// <param name="value">A string to encode.</param>
        public DaitchMokotoffSoundexEncoder(string value)
            : base(value)
        {
        }

        /// <inheritdoc/>
        protected override string EncodeSingleWord(string word)
        {
            if (string.IsNullOrEmpty(word))
            {
                return new string('0', 6);
            }

            var result = new List<string>
            {
                string.Empty,
            };
            LoopThroughString(Transliteration.ToLatin(word), result);

            // Filter out consecutive letters and '_' placeholders.
            return GetFilteredAndPaddedResult(result);
        }

        private static void LoopThroughString(string word, List<string> result)
        {
            var position = 0;
            while (position < word.Length)
            {
                // Iterate through combinations by first letter, check for possible substrings.
                position = UpdatePosition(word, position, result);
            }
        }

        private static string GetFilteredAndPaddedResult(List<string> result)
        {
            for (var index = 0; index < result.Count; index++)
            {
                var cleanedUpValue = GetCleanedUpValue(result[index]);
                result[index] = cleanedUpValue.Length < 6
                    ? cleanedUpValue.PadRight(6, '0')
                    : cleanedUpValue.Substring(0, 6);
            }

            return result.Count == 0 ? string.Empty : result[0];
        }

        private static string GetCleanedUpValue(string codedValue)
        {
            return new string(TrimRepeatingIndexes(codedValue.ToCharArray().ToList()).Where(item => item != '_').ToArray());
        }

        private static int UpdatePosition(string word, int position, List<string> result)
        {
            if (!CombinationsByFirstLetter.ContainsKey(word[position]))
            {
                return position;
            }

            foreach (var combination in CombinationsByFirstLetter[word[position]])
            {
                var newPosition = GetPositionFromSingleCombination(word, position, result, combination);
                if (newPosition == position)
                {
                    continue;
                }

                return newPosition;
            }

            return position;
        }

        private static int GetPositionFromSingleCombination(string word, int position, List<string> result, string combination)
        {
            if (!word.Substring(position).StartsWith(combination))
            {
                // Go to next combination.
                return position;
            }

            string code = GetCombinationCode(word, position, combination);
            if (code.Contains("|"))
            {
                ResolveAmbiguousCode(result, code);
            }
            else
            {
                for (var index = 0; index < result.Count; index++)
                {
                    result[index] = string.Concat(result[index], code);
                }
            }

            return position + combination.Length;
        }

        private static void ResolveAmbiguousCode(List<string> result, string code)
        {
            var splitCode = code.Split('|');
            string[] copy = new string[result.Count];
            for (var index = 0; index < result.Count; index++)
            {
                result[index] = string.Concat(result[index], splitCode[0]);
                copy[index] = string.Concat(result[index], splitCode[1]);
            }

            result.AddRange(copy);
        }

        private static string GetCombinationCode(string word, int position, string combination)
        {
            var codes = CodingChart[combination];

            // Determine positional variant.
            if (position == 0)
            {
                return codes[0];
            }

            if (position + combination.Length < word.Length &&
                     CharsToRemove.Contains(word[position + combination.Length]))
            {
                return codes[1];
            }

            return codes[2];
        }
    }
}
