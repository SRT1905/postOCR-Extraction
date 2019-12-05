using System;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    class MatchProcessor
    {
        public string result;
        public MatchProcessor(string text_to_check, Regex regex_obj, string value_type)
        {
            ProcessValue(text_to_check, regex_obj, value_type);
        }

        private void ProcessValue(string text_to_check, Regex regex_obj, string value_type)
        {
            if (regex_obj.IsMatch(text_to_check))
            {
                MatchCollection matches = regex_obj.Matches(text_to_check);
                result = string.Empty;
                switch (value_type)
                {
                    case "String":
                        result = ProcessString(matches);
                        break;
                    case "Number":
                        result = ProcessNumber(matches);
                        break;
                    case "Date":
                        result = ProcessDate(matches);
                        break;
                }
            }
        }

        private string ProcessString(MatchCollection matches)
        {
            return matches[0].Groups.Count > 1 ? 
                matches[0].Groups[1].Value :
                matches[0].Value;
        }
        private string ProcessNumber(MatchCollection matches)
        {
            var processed_number = Utilities.NumberProcessing(matches[0].Value);
            return double.TryParse(processed_number, out double result) ? result.ToString() : string.Empty;
        }
        private string ProcessDate(MatchCollection matches)
        {
            string result = matches[0].Groups[0].Length != 0 ? matches[0].Groups[0].Value : matches[0].Value;
            result = Utilities.DateProcessing(result);
            return DateTime.TryParse(result, out DateTime _) ? result : string.Empty;
        }
    }
}
