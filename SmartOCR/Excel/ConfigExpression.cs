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

            this.InitializeSearchParameters(valueType, this.ParseInput(input));
        }

        /// <summary>
        /// Gets regular expression pattern.
        /// </summary>
        public string RegExPattern { get; private set; }

        /// <summary>
        /// Gets mapping between search parameter name and its value.
        /// </summary>
        public Dictionary<string, int> SearchParameters { get; private set; }

        private static void TryToMergeSplittedPattern(List<string> splittedInput)
        {
            while (!(int.TryParse(splittedInput[1], out _) || string.IsNullOrEmpty(splittedInput[1])))
            {
                MergeSplittedPattern(splittedInput);
            }
        }

        private static void MergeSplittedPattern(List<string> splittedInput)
        {
            splittedInput[0] = $"{splittedInput[0]};{splittedInput[1]}";
            for (int i = 2; i < splittedInput.Count; i++)
            {
                splittedInput[i - 1] = splittedInput[i];
            }

            splittedInput.RemoveAt(splittedInput.Count - 1);
        }

        private static string[] DefineNumericParameterTitles(string valueType)
        {
            string[] parameterTitles = new string[2] { "row", "column" };

            if (!valueType.Contains("Table"))
            {
                parameterTitles[0] = "line_offset";
                parameterTitles[1] = "horizontal_status";
            }

            return parameterTitles;
        }

        private void InitializeSearchParameters(string valueType, List<string> parsedInput)
        {
            string[] parameterTitles = DefineNumericParameterTitles(valueType);

            this.SearchParameters = new Dictionary<string, int>();
            for (int i = 0; i < 2; i++)
            {
                this.SearchParameters.Add(parameterTitles[i], int.Parse(parsedInput[i + 1], NumberStyles.Integer, NumberFormatInfo.InvariantInfo));
            }
        }

        private List<string> ParseInput(string input)
        {
            if (input == null)
            {
                return new List<string>() { null, "0", "0" };
            }

            List<string> splittedInput = input.Split(';').ToList();
            TryToMergeSplittedPattern(splittedInput);
            this.ValidateNumericParameters(splittedInput);
            return splittedInput;
        }

        private void ValidateNumericParameters(List<string> splittedInput)
        {
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
        }
    }
}
