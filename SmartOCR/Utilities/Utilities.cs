using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SmartOCR
{
    static class Utilities
    {
        public static bool TryAddValueInCollection(long value, ref List<long> value_collection)
        {
            if (value_collection.Contains(value))
                return false;
            value_collection.Add(value);
            return true;
        }

        public static Regex CreateRegexpObject(string text_pattern, bool is_global = true, bool ignore_case = true)
        {
            RegexOptions options = RegexOptions.ECMAScript;
            if (is_global)
            {
                options |= RegexOptions.Multiline;
            }
            if (ignore_case)
            {
                options |= RegexOptions.IgnoreCase;
            }
            Regex regex_obj = new Regex(text_pattern, options);
            return regex_obj;
        }

        public static string NumberProcessing(string value, bool negate = false)
        {
            string regional_decimal_separator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            value = TrimNonNumericChars(value);
            value = value.Contains(",") && value.Contains(".") ?
                        value.Replace(",", string.Empty) :
                        value.Replace(",", regional_decimal_separator);

            value = value.Replace(".", regional_decimal_separator);

            if (double.TryParse(value, out double numeric_value))
            {
                return negate ? (-numeric_value).ToString() : numeric_value.ToString();
            }

            string temp_value = new string(value.ToCharArray()
                                                .Where(c => !Char.IsWhiteSpace(c))
                                                .ToArray());
            if (temp_value.Count(single_char => single_char == '.' || single_char == ',') == 1)
            {
                temp_value = temp_value.Replace(",", regional_decimal_separator);
                temp_value = temp_value.Replace(".", regional_decimal_separator);
                if (double.TryParse(temp_value, out numeric_value))
                {
                    return negate ? (-numeric_value).ToString() : numeric_value.ToString();
                }

                return string.Empty;
            }
            if (double.TryParse(temp_value, out numeric_value))
            {
                return temp_value;
            }

            return value;
        }

        private static string TrimNonNumericChars(string value)
        {
            string pattern = @"[\d\.\,\s\\]";
            string result = string.Empty;
            RegexOptions options = RegexOptions.Multiline;
            Regex regex = new Regex(pattern, options);
            if (regex.IsMatch(value))
            {
                MatchCollection matches = regex.Matches(value);
                foreach (Match item in matches)
                {
                    result += item.Value;
                }
                return result.Trim().Replace('-', ',');
            }
            return value;
        }

        public static string DateProcessing(string value)
        {
            Regex regex_num_date = CreateRegexpObject(@"(\d{2,4})");
            Regex regex_alpha_month = CreateRegexpObject(@"\W?(\d{2})\W?\s*([а-яёa-z]{3,8})\s*(\d{4})");

            if (regex_num_date.IsMatch(value))
            {
                MatchCollection matches = regex_num_date.Matches(value);
                if (matches.Count >= 3)
                {
                    string date_value = $"{matches[1].Value}/{matches[0].Value}/{matches[2].Value}";
                    return DateTime.TryParse(date_value, out DateTime parsed_date)
                        ? parsed_date.ToShortDateString()
                        : $"Дата распознана как {date_value}";
                }
            }

            if (regex_alpha_month.IsMatch(value))
            {
                MatchCollection matches = regex_alpha_month.Matches(value);
                GroupCollection groups = matches[0].Groups;
                string month = ReturnNumericMonth(groups[1].Value.ToLower());
                string date_value = $"{month}/{groups[0].Value}/{groups[2].Value}";
                return DateTime.TryParse(date_value, out DateTime parsed_date)
                    ? parsed_date.ToShortDateString()
                    : $"Дата распознана как {date_value}";
            }

            return "Дата документа не распознана";

        }

        private static string ReturnNumericMonth(string value)
        {
            Dictionary<string, string> month_mapping = new Dictionary<string, string>()
            {
                {"январ", "01" }, {"jan", "01" },
                {"февраля", "02" }, {"feb", "02" },
                {"март", "03" }, {"mar", "03" },
                {"апрел", "04" }, {"apr", "04" },
                {"мая", "05" }, {"май", "05" }, {"may", "05" },
                {"июн", "06" }, {"jun", "06" },
                {"июл", "07" }, {"jul", "07" },
                {"август", "08" }, {"aug", "08" },
                {"сентябр", "09" }, {"sep", "09" },
                {"октябр", "10" }, {"oct", "10" },
                {"ноябр", "11" }, {"nov", "11" },
                {"декабр", "12" }, {"dec", "12" }
            };

            foreach (string key in month_mapping.Keys)
            {
                if (value.Contains(key))
                    return month_mapping[key];
            }
            return "00";
        }
    }
}
