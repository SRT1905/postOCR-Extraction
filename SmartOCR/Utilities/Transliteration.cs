namespace SmartOCR.Utilities
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Used to transliterate word from Cyrillic letters to Latin and vice versa.
    /// </summary>
    public static class Transliteration
    {
        private static readonly Dictionary<string, string> TransliterationDictionary
            = new Dictionary<string, string>
        {
            ["А"] = "A", ["Б"] = "B", ["В"] = "V", ["Г"] = "G",
            ["Д"] = "D", ["Е"] = "E", ["Ё"] = "JO", ["Ж"] = "ZH",
            ["З"] = "Z", ["И"] = "I", ["Й"] = "J", ["К"] = "K",
            ["Л"] = "L", ["М"] = "M", ["Н"] = "N", ["О"] = "O",
            ["П"] = "P", ["Р"] = "R", ["С"] = "S",  ["Т"] = "T",
            ["У"] = "U", ["Ф"] = "F", ["Х"] = "H", ["Ц"] = "Z",
            ["Ч"] = "CH", ["Ш"] = "SH", ["Щ"] = "SCH", ["Ъ"] = string.Empty,
            ["Ы"] = "Y", ["Ь"] = string.Empty, ["Э"] = "E", ["Ю"] = "JU",
            ["Я"] = "JA",
        };

        /// <summary>
        /// Transliterates Cyrillic letters to Latin.
        /// </summary>
        /// <param name="word">A word to transform.</param>
        /// <param name="toLatin">Indicates whether word should be transliterated to Latin.</param>
        /// <returns>A transliterated word.</returns>
        public static string TransliterateWord(string word, bool toLatin = true)
        {
            return TransliterationDictionary.Aggregate(word, (current, translitPair) => toLatin
                                                                                            ? current.Replace(translitPair.Key, translitPair.Value)
                                                                                            : current.Replace(translitPair.Value, translitPair.Key));
        }
    }
}
