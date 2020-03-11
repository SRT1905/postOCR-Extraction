namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Provides static methods used accross namespace.
    /// </summary>
    public static class Utilities
    {
        private static readonly Dictionary<string, string> MonthMapping = new Dictionary<string, string>()
        {
            { "январ", "01" }, { "jan", "01" },
            { "февраля", "02" }, { "feb", "02" },
            { "март", "03" }, { "mar", "03" },
            { "апрел", "04" }, { "apr", "04" },
            { "мая", "05" }, { "май", "05" }, { "may", "05" },
            { "июн", "06" }, { "jun", "06" },
            { "июл", "07" }, { "jul", "07" },
            { "август", "08" }, { "aug", "08" },
            { "сентябр", "09" }, { "sep", "09" },
            { "октябр", "10" }, { "oct", "10" },
            { "ноябр", "11" }, { "nov", "11" },
            { "декабр", "12" }, { "dec", "12" },
        };

        private static readonly string RegionalDecimalSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        private static readonly Regex RegexAlphaMonth = CreateRegexpObject(@"\W?(\d{2})\W?\s*([а-яёa-z]{3,8})\s*(\d{4})");
        private static readonly Regex RegexNumDate = CreateRegexpObject(@"(\d{2,4})");

        /// <summary>
        /// Gets a value indicating whether debug logging is enabled during current run.
        /// </summary>
        public static bool EnableDebug { get; internal set; }

        /// <summary>
        /// Gets or sets a lowest level of showing debug messages.
        /// </summary>
        public static int DebugLevel { get; set; } = -1;

        /// <summary>
        /// Creates a new <see cref="Regex"/> instance with multiline and ignore case mode.
        /// </summary>
        /// <param name="textPattern">Search pattern.</param>
        /// <returns><see cref="Regex"/> object.</returns>
        public static Regex CreateRegexpObject(string textPattern)
        {
            return new Regex(textPattern, RegexOptions.ECMAScript | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Sends debug message to <see cref="Console.Out"/>.
        /// </summary>
        /// <param name="message">A message string.</param>
        /// <param name="debugLevel">Used to indicate call level.</param>
        public static void Debug(string message, int debugLevel = 0)
        {
            if (Utilities.EnableDebug)
            {
                if (debugLevel <= Utilities.DebugLevel || Utilities.DebugLevel == -1)
                {
                    Console.Write($"# {new string(' ', debugLevel)}");
                    Console.WriteLine(message);
                }
            }
        }

        /// <summary>
        /// Tries to process passed string as date.
        /// </summary>
        /// <param name="value">String, containing date.</param>
        /// <returns>String representation of date value.</returns>
        public static string DateProcessing(string value)
        {
            return ProcessDateAsNumeric(value) ?? ProcessDateAsAlphaNumeric(value) ?? "Дата документа не распознана";
        }

        /// <summary>
        /// Tries to represent passed string as number.
        /// </summary>
        /// <param name="value">String, containing number.</param>
        /// <returns>String representation of parsed number.</returns>
        public static string NumberProcessing(string value)
        {
            value = PrepareValueForNumberProcessing(value);
            return double.TryParse(value, out double _)
                ? value
                : value.Count(singleChar => singleChar == '.' || singleChar == ',') == 1
                    ? GetNumberAfterRemovingSeparators(value)
                    : value;
        }

        private static string GetNumberAfterRemovingSeparators(string value)
        {
            value = value.Replace(",", RegionalDecimalSeparator)
                         .Replace(".", RegionalDecimalSeparator);
            return double.TryParse(value, out _)
                ? value
                : string.Empty;
        }

        private static string PrepareValueForNumberProcessing(string value)
        {
            value = ProcessSeparatorsInNumber(TrimNonNumericChars(value));
            return new string(value.ToCharArray()
                                   .Where(c => !char.IsWhiteSpace(c))
                                   .ToArray());
        }

        private static string ProcessDateAsAlphaNumeric(string value)
        {
            if (!RegexAlphaMonth.IsMatch(value))
            {
                return null;
            }

            GroupCollection groups = RegexAlphaMonth.Matches(value)[0].Groups;
            string month = ReturnNumericMonth(groups[2].Value.ToLower());
            return TryReturnDateValue($"{month}/{groups[1].Value}/{groups[3].Value}");
        }

        private static string ProcessDateAsNumeric(string value)
        {
            if (RegexNumDate.IsMatch(value))
            {
                MatchCollection matches = RegexNumDate.Matches(value);
                if (matches.Count >= 3)
                {
                    return TryReturnDateValue($"{matches[1].Value}/{matches[0].Value}/{matches[2].Value}");
                }
            }

            return null;
        }

        private static string ProcessSeparatorsInNumber(string value)
        {
            value = value.Contains(",") && value.Contains(".")
                ? value.Replace(",", string.Empty)
                : value.Replace(",", RegionalDecimalSeparator);

            value = value.Replace(".", RegionalDecimalSeparator);
            return value;
        }

        private static string ReturnNumericMonth(string value)
        {
            foreach (var item in MonthMapping)
            {
                if (value.Contains(item.Key))
                {
                    return item.Value;
                }
            }

            return "00";
        }

        private static string TrimNonNumericChars(string value)
        {
            Regex regexObject = new Regex(@"[\d\.\,\s\\]", RegexOptions.Multiline);
            return regexObject.IsMatch(value)
                ? string.Join(
                    string.Empty,
                    regexObject.Matches(value)
                               .Cast<Match>()
                               .Select(match => match.Value)).Trim().Replace('-', ',')
                : value;
        }

        private static string TryReturnDateValue(string dateValue)
        {
            return DateTime.TryParse(dateValue, out DateTime _)
                ? dateValue
                : $"Дата распознана как {dateValue}";
        }
    }
}