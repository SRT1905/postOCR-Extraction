namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;

    /// <summary>
    /// Describes single search expression defined in Excel config file.
    /// </summary>
    public class ConfigExpression
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigExpression"/> class.
        /// Instance is initialized by Excel cell contents.
        /// </summary>
        /// <param name="input">Excel cell contents, containing regular expression pattern, line offset and horizontal search status.</param>
        /// <param name="valueType">String representation of field data type.</param>
        public ConfigExpression(string valueType, string input)
        {
            if (string.IsNullOrEmpty(valueType))
            {
                throw new ArgumentNullException(nameof(valueType));
            }

            List<string> parsedInput = this.ParseInput(input);
            if (valueType.Contains("Table"))
            {
                this.SearchParameters = new Dictionary<string, int>()
                {
                    { "row", int.Parse(parsedInput[1], NumberStyles.Integer, NumberFormatInfo.InvariantInfo) },
                    { "column", int.Parse(parsedInput[2], NumberStyles.Integer, NumberFormatInfo.InvariantInfo) },
                };
            }
            else
            {
                this.SearchParameters = new Dictionary<string, int>()
                {
                    { "line_offset", int.Parse(parsedInput[1], NumberStyles.Integer, NumberFormatInfo.InvariantInfo) },
                    { "horizontal_status", int.Parse(parsedInput[2], NumberStyles.Integer, NumberFormatInfo.InvariantInfo) },
                };
            }
        }

        public string RegExPattern { get; private set; }

        public Dictionary<string, int> SearchParameters { get; }

        private List<string> ParseInput(string input)
        {
            if (input == null)
            {
                return new List<string>() { null, "0", "0" };
            }

            List<string> splittedInput = input.Split(';').ToList();
            while (!(int.TryParse(splittedInput[1], out _) || string.IsNullOrEmpty(splittedInput[1])))
            {
                splittedInput[0] = $"{splittedInput[0]};{splittedInput[1]}";
                for (int i = 2; i < splittedInput.Count; i++)
                {
                    splittedInput[i - 1] = splittedInput[i];
                }

                splittedInput.RemoveAt(splittedInput.Count - 1);
            }

            while (splittedInput.Count < 3)
            {
                splittedInput.Add("0");
            }

            this.RegExPattern = splittedInput[0];
            for (int i = 1; i < splittedInput.Count; i++)
            {
                if (string.IsNullOrEmpty(splittedInput[i]))
                {
                    splittedInput[i] = "0";
                }
            }

            return splittedInput;
        }
    }
}
