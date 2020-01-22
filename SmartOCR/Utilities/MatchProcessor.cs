using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    /// <summary>
    /// Used to process regular expression match as distinct data type value.
    /// </summary>
    internal class MatchProcessor
    {
        /// <summary>
        /// Match, processed as distinct datatype value.
        /// </summary>
        public string Result { get; }

        /// <summary>
        /// Initializes a new <see cref="MatchProcessor"/> instance that takes data type, <see cref="Regex"/> object and text to check.
        /// </summary>
        /// <param name="regex_obj"><see cref="Regex"/> object that is used to test paragraphs.</param>
        /// <param name="value_type">Data type of found matches.</param>
        public MatchProcessor(string text_to_check, Regex regex_obj, string value_type)
        {
            Result = ProcessValue(text_to_check, regex_obj, value_type);
        }

        /// <summary>
        /// Calls for specific type processing for possible match between text and regular expression.
        /// </summary>
        /// <param name="regex_obj"><see cref="Regex"/> object that is used to test paragraphs.</param>
        /// <param name="value_type">Data type of found matches.</param>
        /// <returns>String representation of casted match.</returns>
        private string ProcessValue(string text_to_check, Regex regex_obj, string value_type)
        {
            if (regex_obj.IsMatch(text_to_check))
            {
                switch (value_type)
                {
                    case "String":
                        return ProcessString(regex_obj.Matches(text_to_check));

                    case "Number":
                        return ProcessNumber(regex_obj.Matches(text_to_check));

                    case "Date":
                        return ProcessDate(regex_obj.Matches(text_to_check));

                    default:
                        break;
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Gets first match in collection. Return capturing group if there is one.
        /// </summary>
        /// <param name="matches">Collection of found matches.</param>
        /// <returns>String representation of match.</returns>
        private string ProcessString(MatchCollection matches)
        {
            return matches[0].Groups.Count > 1 
                ? matches[0].Groups[1].Value 
                : matches[0].Value;
        }

        /// <summary>
        /// Tries to parse match as number.
        /// </summary>
        /// <param name="matches">Collection of found matches.</param>
        /// <returns>String representation of numeric match.</returns>
        private string ProcessNumber(MatchCollection matches)
        {
            var processed_number = Utilities.NumberProcessing(matches[0].Value);
            return double.TryParse(processed_number, out double result) 
                ? result.ToString(CultureInfo.CurrentCulture) 
                : string.Empty;
        }

        /// <summary>
        /// Tries to parse match as date.
        /// </summary>
        /// <param name="matches">Collection of found matches.</param>
        /// <returns>String representation of date match.</returns>
        private string ProcessDate(MatchCollection matches)
        {
            string result = matches[0].Groups[1].Length != 0 
                ? matches[0].Groups[1].Value 
                : matches[0].Value;
            result = Utilities.DateProcessing(result);
            return DateTime.TryParse(result, out DateTime _) 
                ? result 
                : string.Empty;
        }
    }
}