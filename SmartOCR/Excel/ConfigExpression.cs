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
            OffsetInputByOneItemToLeft(splittedInput);

            splittedInput.RemoveAt(splittedInput.Count - 1);
        }

        private static void OffsetInputByOneItemToLeft(List<string> splittedInput)
        {
            for (int i = 2; i < splittedInput.Count; i++)
            {
                splittedInput[i - 1] = splittedInput[i];
            }
        }

        private static string[] DefineNumericParameterTitles(string valueType)
        {
            string[] parameterTitles = new string[2] { "row", "column" };
            string[] tableParameterTitles = new string[2] { "line_offset", "horizontal_status" };

            return valueType.Contains("Table")
                ? tableParameterTitles
                : parameterTitles;
        }

        private static Dictionary<string, int> MapParametersWithValues(List<string> parsedInput, string[] parameterTitles)
        {
            return new Dictionary<string, int>()
            {
                { parameterTitles[0], int.Parse(parsedInput[1], NumberStyles.Integer, NumberFormatInfo.InvariantInfo) },
                { parameterTitles[1], int.Parse(parsedInput[2], NumberStyles.Integer, NumberFormatInfo.InvariantInfo) },
            };
        }

        private static void AddZerosToEnd(List<string> splittedInput)
        {
            while (splittedInput.Count < 3)
            {
                splittedInput.Add("0");
            }
        }

        private static void TrySetDefaultNumericValues(List<string> splittedInput)
        {
            for (int i = 1; i < splittedInput.Count; i++)
            {
                if (string.IsNullOrEmpty(splittedInput[i]))
                {
                    splittedInput[i] = "0";
                }
            }
        }

        private void InitializeSearchParameters(string valueType, List<string> parsedInput)
        {
            string[] parameterTitles = DefineNumericParameterTitles(valueType);
            this.SearchParameters = MapParametersWithValues(parsedInput, parameterTitles);
        }

        private List<string> ParseInput(string input)
        {
            if (input == null)
            {
                return new List<string>() { null, "0", "0" };
            }

            return this.DefineExpressionParameters(input);
        }

        private List<string> DefineExpressionParameters(string input)
        {
            List<string> splittedInput = input.Split(';').ToList();
            TryToMergeSplittedPattern(splittedInput);
            return this.ValidateNumericParameters(splittedInput);
        }

        private List<string> ValidateNumericParameters(List<string> splittedInput)
        {
            AddZerosToEnd(splittedInput);

            this.RegExPattern = splittedInput[0];
            TrySetDefaultNumericValues(splittedInput);

            return splittedInput;
        }
    }
}
