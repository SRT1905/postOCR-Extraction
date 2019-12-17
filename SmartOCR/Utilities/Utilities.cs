using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SmartOCR
{
    /// <summary>
    /// Provides static methods used accross namespace.
    /// </summary>
    internal static class Utilities
    {
        /// <summary>
        /// Contains names of document types that are supported for processing.
        /// </summary>
        public static HashSet<string> valid_document_types = new HashSet<string>
            {
                "invoice_sales",
                "settlement",
                "invoice_costs",
            };

        /// <summary>
        /// Creates a new <see cref="Regex"/> instance with additional parameters.
        /// </summary>
        /// <param name="text_pattern">Search pattern.</param>
        /// <param name="is_multiline">Indicates that characters '^' and '$' will match beginning and end, respectively, of any line.</param>
        /// <param name="ignore_case">Indicates that expression will be case-insensitive.</param>
        /// <returns><see cref="Regex"/> object.</returns>
        public static Regex CreateRegexpObject(string text_pattern, bool is_multiline = true, bool ignore_case = true)
        {
            RegexOptions options = RegexOptions.ECMAScript;
            if (is_multiline)
            {
                options |= RegexOptions.Multiline;
            }
            if (ignore_case)
            {
                options |= RegexOptions.IgnoreCase;
            }
            return new Regex(text_pattern, options);
        }

        /// <summary>
        /// Tries to represent passed string as number.
        /// </summary>
        /// <param name="value">String, containing number.</param>
        /// <returns>String representation of parsed number.</returns>
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

        /// <summary>
        /// Removes non-numeric characters and replaces decimal separator.
        /// </summary>
        /// <param name="value">String, containing number.</param>
        /// <param name="regional_decimal_separator">Character that is set by regional settings to separate integer and fraction parts of number.</param>
        /// <returns>String, prepared for parsing.</returns>
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

        /// <summary>
        /// Character that is set by regional settings to separate integer and fraction parts of number.
        /// </summary>
        /// <returns>String, containing regional decimal separator.</returns>
        private static string GetDecimalSeparator()
        {
            return Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        }

        /// <summary>
        /// Trims string of characters that cannot be part of number.
        /// </summary>
        /// <param name="value">String, containing number.</param>
        /// <returns>String, trimmed of invalid characters.</returns>
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

        /// <summary>
        /// Tries to process passed string as date.
        /// </summary>
        /// <param name="value">String, containing date.</param>
        /// <returns>String representation of date value.</returns>
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

        /// <summary>
        /// Searches for numerical representation of month in passed string.
        /// </summary>
        /// <param name="value">String, containing month.</param>
        /// <returns>String representation of numerical month.</returns>
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

            List<string> keys = month_mapping.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                if (value.Contains(key))
                {
                    return month_mapping[key];
                }
            }
            return "00";
        }

        /// <summary>
        /// Used to print message to command prompt.
        /// </summary>
        public static void PrintInvalidInputMessage()
        {
            Console.WriteLine("Enter valid document type and path(s) to file/directory.");
            Thread.Sleep(2000);
        }
    }
}