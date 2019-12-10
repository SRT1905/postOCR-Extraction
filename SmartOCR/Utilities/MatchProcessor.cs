using System;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    internal class MatchProcessor
    {
        public string result { get; }

        public MatchProcessor(string text_to_check, Regex regex_obj, string value_type)
        {
            result = ProcessValue(text_to_check, regex_obj, value_type);
        }

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

        private string ProcessString(MatchCollection matches)
        {
            return matches[0].Groups.Count > 1 
                ? matches[0].Groups[1].Value 
                : matches[0].Value;
        }

        private string ProcessNumber(MatchCollection matches)
        {
            var processed_number = Utilities.NumberProcessing(matches[0].Value);
            return double.TryParse(processed_number, out double result) ? result.ToString() : string.Empty;
        }

        private string ProcessDate(MatchCollection matches)
        {
            string result = matches[0].Groups[1].Length != 0 ? matches[0].Groups[1].Value : matches[0].Value;
            result = Utilities.DateProcessing(result);
            return DateTime.TryParse(result, out DateTime _) ? result : string.Empty;
        }
    }
}