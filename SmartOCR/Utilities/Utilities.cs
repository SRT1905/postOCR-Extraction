﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SmartOCR
{
    /// <summary>
    /// Provides static methods used accross namespace.
    /// </summary>
    public static class Utilities
    {
        #region Public static methods
        /// <summary>
        /// Creates a new <see cref="Regex"/> instance with additional parameters.
        /// </summary>
        /// <param name="textPattern">Search pattern.</param>
        /// <param name="isMultiline">Indicates that characters '^' and '$' will match beginning and end, respectively, of any line.</param>
        /// <param name="ignoreCase">Indicates that expression will be case-insensitive.</param>
        /// <returns><see cref="Regex"/> object.</returns>
        public static Regex CreateRegexpObject(string textPattern, bool isMultiline = true, bool ignoreCase = true)
        {
            RegexOptions options = RegexOptions.ECMAScript;
            if (isMultiline)
            {
                options |= RegexOptions.Multiline;
            }
            if (ignoreCase)
            {
                options |= RegexOptions.IgnoreCase;
            }
            return new Regex(textPattern, options);
        }
        public static void Debug(string format, int debugLevel = 0, params object[] args)
        {
            Console.Write(Properties.Resources.debugHashtag);
            Console.Write(new string('\t', debugLevel));
            Console.WriteLine(format, args);
        }
        /// <summary>
        /// Tries to process passed string as date.
        /// </summary>
        /// <param name="value">String, containing date.</param>
        /// <returns>String representation of date value.</returns>
        public static string DateProcessing(string value)
        {
            Regex regexNumDate = CreateRegexpObject(@"(\d{2,4})");
            Regex regexAlphaMonth = CreateRegexpObject(@"\W?(\d{2})\W?\s*([а-яёa-z]{3,8})\s*(\d{4})");

            if (regexNumDate.IsMatch(value))
            {
                MatchCollection matches = regexNumDate.Matches(value);
                if (matches.Count >= 3)
                {
                    string dateValue = $"{matches[1].Value}/{matches[0].Value}/{matches[2].Value}";
                    return DateTime.TryParse(dateValue, out DateTime _)
                        ? dateValue
                        : $"Дата распознана как {dateValue}";
                }
            }

            if (regexAlphaMonth.IsMatch(value))
            {
                GroupCollection groups = regexAlphaMonth.Matches(value)[0].Groups;
                string month = ReturnNumericMonth(groups[2].Value.ToLower(CultureInfo.CurrentCulture));
                string dateValue = $"{month}/{groups[1].Value}/{groups[3].Value}";
                return DateTime.TryParse(dateValue, out DateTime _)
                    ? dateValue
                    : $"Дата распознана как {dateValue}";
            }

            return "Дата документа не распознана";
        }
        /// <summary>
        /// Tries to represent passed string as number.
        /// </summary>
        /// <param name="value">String, containing number.</param>
        /// <returns>String representation of parsed number.</returns>
        public static string NumberProcessing(string value)
        {
            string regionalDecimalSeparator = GetDecimalSeparator();
            value = PrepareValueForNumberProcessing(value, regionalDecimalSeparator);

            if (double.TryParse(value, out double _))
            {
                return value;
            }

            if (value.Count(singleChar => singleChar == '.' || singleChar == ',') == 1)
            {
                value = value.Replace(",", regionalDecimalSeparator)
                             .Replace(".", regionalDecimalSeparator);
                return double.TryParse(value, out _) ? value : string.Empty;
            }
            return value;
        }
        /// <summary>
        /// Used to print message to command prompt.
        /// </summary>
        public static void PrintInvalidInputMessage()
        {
            Console.WriteLine(Properties.Resources.invalidInputMessage);
            Thread.Sleep(2000);
        }
        #endregion

        #region Private static methods
        /// <summary>
        /// Character that is set by regional settings to separate integer and fraction parts of number.
        /// </summary>
        /// <returns>String, containing regional decimal separator.</returns>
        private static string GetDecimalSeparator()
        {
            return Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        }
        /// <summary>
        /// Removes non-numeric characters and replaces decimal separator.
        /// </summary>
        /// <param name="value">String, containing number.</param>
        /// <param name="regionalDecimalSeparator">Character that is set by regional settings to separate integer and fraction parts of number.</param>
        /// <returns>String, prepared for parsing.</returns>
        private static string PrepareValueForNumberProcessing(string value, string regionalDecimalSeparator)
        {
            value = TrimNonNumericChars(value);
            value = value.Contains(",") && value.Contains(".")
                ? value.Replace(",", string.Empty)
                : value.Replace(",", regionalDecimalSeparator);

            value = value.Replace(".", regionalDecimalSeparator);
            return new string(value.ToCharArray()
                                   .Where(c => !char.IsWhiteSpace(c))
                                   .ToArray());
        }
        /// <summary>
        /// Searches for numerical representation of month in passed string.
        /// </summary>
        /// <param name="value">String, containing month.</param>
        /// <returns>String representation of numerical month.</returns>
        private static string ReturnNumericMonth(string value)
        {
            Dictionary<string, string> monthMapping = new Dictionary<string, string>()
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

            List<string> keys = monthMapping.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                if (value.Contains(key))
                {
                    return monthMapping[key];
                }
            }
            return "00";
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
        #endregion
    }
}