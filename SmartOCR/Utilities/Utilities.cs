using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SmartOCR
{
    internal static class Utilities
    {
        public static HashSet<string> valid_document_types = new HashSet<string>
            {
                "Invoice"
            };
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

        public static string NumberProcessing(string value)
        {
            string regional_decimal_separator = GetDecimalSeparator();
            value = PrepareValueForNumberProcessing(value, regional_decimal_separator);

            if (double.TryParse(value, out double _))
            {
                return value;
            }

            if (value.Count(single_char => single_char == '.' || single_char == ',') == 1)
            {
                value = value.Replace(",", regional_decimal_separator)
                             .Replace(".", regional_decimal_separator);
                return double.TryParse(value, out _) ? value : string.Empty;
            }
            return value;
        }

        private static string PrepareValueForNumberProcessing(string value, string regional_decimal_separator)
        {
            value = TrimNonNumericChars(value);
            value = value.Contains(",") && value.Contains(".")
                ? value.Replace(",", string.Empty)
                : value.Replace(",", regional_decimal_separator);

            value = value.Replace(".", regional_decimal_separator);
            return new string(value.ToCharArray()
                                   .Where(c => !char.IsWhiteSpace(c))
                                   .ToArray());
        }

        private static string GetDecimalSeparator()
        {
            return Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        }

        private static string TrimNonNumericChars(string value)
        {
            Regex regex = new Regex(@"[\d\.\,\s\\]", RegexOptions.Multiline);
            if (regex.IsMatch(value))
            {
                MatchCollection matches = regex.Matches(value);
                var result = string.Join("", matches.Cast<Match>().Select(match => match.Value));
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
                    return DateTime.TryParse(date_value, out DateTime _)
                        ? date_value
                        : $"Дата распознана как {date_value}";
                }
            }

            if (regex_alpha_month.IsMatch(value))
            {
                GroupCollection groups = regex_alpha_month.Matches(value)[0].Groups;
                string month = ReturnNumericMonth(groups[2].Value.ToLower());
                string date_value = $"{month}/{groups[1].Value}/{groups[3].Value}";
                return DateTime.TryParse(date_value, out DateTime _)
                    ? date_value
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
                {
                    return month_mapping[key];
                }
            }
            return "00";
        }
    }
}