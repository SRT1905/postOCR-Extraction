namespace SmartOCR.Config
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using SmartOCR.Soundex;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

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
        /// Gets text pattern (Soundex or Regex).
        /// </summary>
        public string TextExpression { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the value of <see cref="TextExpression"/> is Soundex expression.
        /// </summary>
        public bool UseSoundex { get; private set; }

        /// <summary>
        /// Gets mapping between search parameter name and its value.
        /// </summary>
        public Dictionary<string, int> SearchParameters { get; private set; }

        private static void TryToMergeSplitPattern(List<string> splitInput)
        {
            while (!(int.TryParse(splitInput[1], out _) || string.IsNullOrEmpty(splitInput[1])))
            {
                MergeSplitPattern(splitInput);
            }
        }

        private static void MergeSplitPattern(List<string> splitInput)
        {
            splitInput[0] = $"{splitInput[0]};{splitInput[1]}";
            OffsetInputByOneItemToLeft(splitInput);
            splitInput.RemoveAt(splitInput.Count - 1);
        }

        private static void OffsetInputByOneItemToLeft(List<string> splitInput)
        {
            for (int i = 2; i < splitInput.Count; i++)
            {
                splitInput[i - 1] = splitInput[i];
            }
        }

        private static string[] DefineNumericParameterTitles(string valueType)
        {
            return valueType.Contains("Table")
                ? new[] { "row", "column" }
                : new[] { "line_offset", "horizontal_status" };
        }

        private static Dictionary<string, int> MapParametersWithValues(List<string> parsedInput, string[] parameterTitles)
        {
            return new Dictionary<string, int>
            {
                [parameterTitles[0]] = int.Parse(parsedInput[1], NumberStyles.Integer, NumberFormatInfo.InvariantInfo),
                [parameterTitles[1]] = int.Parse(parsedInput[2], NumberStyles.Integer, NumberFormatInfo.InvariantInfo),
            };
        }

        private static void AddZerosToEnd(List<string> splitInput)
        {
            while (splitInput.Count < 3)
            {
                splitInput.Add("0");
            }
        }

        private static void TrySetDefaultNumericValues(List<string> splitInput)
        {
            for (int i = 1; i < splitInput.Count; i++)
            {
                if (string.IsNullOrEmpty(splitInput[i]))
                {
                    splitInput[i] = "0";
                }
            }
        }

        private static string EncodeString(string value)
        {
            // return new DefaultSoundexEncoder(value).EncodedValue;
            return new DaitchMokotoffSoundexEncoder(value).EncodedValue;
        }

        private void InitializeSearchParameters(string valueType, List<string> parsedInput)
        {
            string[] parameterTitles = DefineNumericParameterTitles(valueType);
            this.SearchParameters = MapParametersWithValues(parsedInput, parameterTitles);
        }

        private List<string> ParseInput(string input)
        {
            return input == null
                ? new List<string>() { null, "0", "0" }
                : this.DefineExpressionParameters(input);
        }

        private List<string> DefineExpressionParameters(string input)
        {
            List<string> splitInput = input.Split(';').ToList();
            TryToMergeSplitPattern(splitInput);
            return this.ValidateNumericParameters(splitInput);
        }

        private List<string> ValidateNumericParameters(List<string> splitInput)
        {
            AddZerosToEnd(splitInput);
            this.TextExpression = splitInput[0];
            this.ValidateSoundexStatus();
            TrySetDefaultNumericValues(splitInput);
            return splitInput;
        }

        private void ValidateSoundexStatus()
        {
            if (this.TextExpression.StartsWith("soundex"))
            {
                this.PopulatePropertiesWithSoundex();
            }
        }

        private void PopulatePropertiesWithSoundex()
        {
            this.TextExpression = EncodeString(this.ExtractSoundexExpression());
            this.UseSoundex = true;
        }

        private string ExtractSoundexExpression()
        {
            return Utilities.CreateRegexpObject(@"soundex\((.*)\)")
                            .Match(this.TextExpression)
                            .Groups[1]
                            .Value;
        }
    }
}
