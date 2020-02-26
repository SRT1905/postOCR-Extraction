namespace SmartOCR
{
    using System;
    using System.Globalization;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to process regular expression match as distinct data type value.
    /// </summary>
    public class MatchProcessor
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MatchProcessor"/> class.
        /// Instance takes data type, <see cref="Regex"/> object and text to check.
        /// </summary>
        /// <param name="textToCheck">Text to check.</param>
        /// <param name="regExObject"><see cref="Regex"/> object that is used to test paragraphs.</param>
        /// <param name="valueType">Data type of found matches.</param>
        public MatchProcessor(string textToCheck, Regex regExObject, string valueType)
        {
            if (regExObject == null)
            {
                this.Result = null;
                return;
            }

            this.Result = this.ProcessValue(textToCheck, regExObject, valueType);
        }

        /// <summary>
        /// Gets match, processed as distinct datatype value.
        /// </summary>
        public string Result { get; }

        /// <summary>
        /// Tries to parse match as date.
        /// </summary>
        /// <param name="matches">Collection of found matches.</param>
        /// <returns>String representation of date match.</returns>
        private static string ProcessDate(MatchCollection matches)
        {
            string result = matches[0].Groups[1].Length != 0
                ? matches[0].Groups[1].Value
                : matches[0].Value;
            result = Utilities.DateProcessing(result);
            return DateTime.TryParse(result, out DateTime _)
                ? result
                : string.Empty;
        }

        /// <summary>
        /// Tries to parse match as number.
        /// </summary>
        /// <param name="matches">Collection of found matches.</param>
        /// <returns>String representation of numeric match.</returns>
        private static string ProcessNumber(MatchCollection matches)
        {
            var processedNumber = Utilities.NumberProcessing(matches[0].Value);
            return double.TryParse(processedNumber, out double result)
                ? result.ToString(CultureInfo.CurrentCulture)
                : string.Empty;
        }

        /// <summary>
        /// Gets first match in collection. Return capturing group if there is one.
        /// </summary>
        /// <param name="matches">Collection of found matches.</param>
        /// <returns>String representation of match.</returns>
        private static string ProcessString(MatchCollection matches)
        {
            return matches[0].Groups.Count > 1
                ? matches[0].Groups[1].Value
                : matches[0].Value;
        }

        /// <summary>
        /// Calls for specific type processing for possible match between text and regular expression.
        /// </summary>
        /// <param name="regexObject"><see cref="Regex"/> object that is used to test paragraphs.</param>
        /// <param name="valueType">Data type of found matches.</param>
        /// <returns>String representation of casted match.</returns>
        private string ProcessValue(string textToCheck, Regex regexObject, string valueType)
        {
            if (regexObject.IsMatch(textToCheck))
            {
                switch (valueType)
                {
                    case "String":
                        return ProcessString(regexObject.Matches(textToCheck));

                    case "Number":
                        return ProcessNumber(regexObject.Matches(textToCheck));

                    case "Date":
                        return ProcessDate(regexObject.Matches(textToCheck));

                    default:
                        break;
                }
            }

            return string.Empty;
        }
    }
}