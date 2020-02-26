using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// Describes single search expression defined in Excel config file.
    /// </summary>
    public class ConfigExpression
    {
        #region Fields
        private readonly Dictionary<string, int> search_parameters;
        #endregion

        #region Properties
        public string RegExPattern { get; private set; }
        public Dictionary<string, int> SearchParameters
        {
            get
            {
                return search_parameters;
            }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new <see cref="ConfigExpression"/> instance with Excel cell contents.
        /// </summary>
        /// <param name="input">Excel cell contents, containing regular expression pattern, line offset and horizontal search status.</param>
        public ConfigExpression(string valueType, string input)
        {
            if (string.IsNullOrEmpty(valueType))
            {
                throw new ArgumentNullException(nameof(valueType));
            }
            List<string> parsed_input = ParseInput(input);
            if (valueType.Contains("Table"))
            {
                search_parameters = new Dictionary<string, int>()
                {
                    {"row", int.Parse(parsed_input[1], NumberStyles.Integer, NumberFormatInfo.InvariantInfo)},
                    {"column", int.Parse(parsed_input[2], NumberStyles.Integer, NumberFormatInfo.InvariantInfo)},
                };
            }
            else
            {
                search_parameters = new Dictionary<string, int>()
                {
                    {"line_offset", int.Parse(parsed_input[1], NumberStyles.Integer, NumberFormatInfo.InvariantInfo)},
                    {"horizontal_status", int.Parse(parsed_input[2], NumberStyles.Integer, NumberFormatInfo.InvariantInfo)},
                };
            }
        }
        private List<string> ParseInput(string input)
        {
            if (input == null)
            {
                return new List<string>() { null, "0", "0" };
            }
            List<string> splitted_input = input.Split(';').ToList();
            while (!(int.TryParse(splitted_input[1], out _) || string.IsNullOrEmpty(splitted_input[1])))
            {
                splitted_input[0] = $"{splitted_input[0]};{splitted_input[1]}";
                for (int i = 2; i < splitted_input.Count; i++)
                {
                    splitted_input[i - 1] = splitted_input[i];
                }
                splitted_input.RemoveAt(splitted_input.Count - 1);
            }
            while (splitted_input.Count < 3)
            {
                splitted_input.Add("0");
            }

            RegExPattern = splitted_input[0];
            for (int i = 1; i < splitted_input.Count; i++)
            {
                if (string.IsNullOrEmpty(splitted_input[i]))
                {
                    splitted_input[i] = "0";
                }
            }
            return splitted_input;
        }
        #endregion
    }
}
