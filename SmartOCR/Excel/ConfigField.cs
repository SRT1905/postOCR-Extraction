namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Describes single search field defined in Excel config file.
    /// </summary>
    public class ConfigField
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigField"/> class.
        /// Instance is initialized with name and value type.
        /// </summary>
        /// <param name="name">Field name.</param>
        /// <param name="valueType">Data type of searched value.</param>
        public ConfigField(string name, string valueType)
        {
            this.Name = name;
            this.ValueType = valueType;
        }

        /// <summary>
        /// Gets collection of search expressions that define search process for current field.
        /// </summary>
        public List<ConfigExpression> Expressions { get; } = new List<ConfigExpression>();

        /// <summary>
        /// Gets or sets string value that is used to check similarity between it and values that match <see cref="TextExpression"/>.
        /// </summary>
        public string ExpectedName { get; set; }

        /// <summary>
        /// Gets or sets identifying expression - whether Regex pattern or Soundex value.
        /// </summary>
        public string TextExpression { get; set; }

        /// <summary>
        /// Gets field name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets data type of searched value.
        /// </summary>
        public string ValueType { get; }

        /// <summary>
        /// Gets a value indicating whether field identifier should be searched using <see cref="ParagraphContainer.Soundex"/> instead of <see cref="ParagraphContainer.Text"/>.
        /// </summary>
        public bool UseSoundex { get; private set; }

        /// <summary>
        /// Gets a value pair, indicating grid coordinates where field search should be performed.
        /// </summary>
        public Tuple<int, int> GridCoordinates { get; private set; }

        /// <summary>
        /// Parses Excel cell contents and gets regular expression pattern and expected field name.
        /// </summary>
        /// <param name="input">String representation of Excel cell contents.</param>
        public void ParseFieldExpression(string input)
        {
            if (input == null)
            {
                this.TextExpression = this.Name;
                this.ExpectedName = this.Name;
                return;
            }

            this.SplitInputAndSetProperties(input);
        }

        /// <summary>
        /// Parses Excel cell contents and gets coordinates of grid structure, where search should be done.
        /// </summary>
        /// <param name="coordinatesValue">String representation of Excel cell contents.</param>
        public void ParseGridCoordinates(string coordinatesValue)
        {
            if (string.IsNullOrEmpty(coordinatesValue))
            {
                this.GridCoordinates = new Tuple<int, int>(-1, -1);
                return;
            }

            this.GridCoordinates = ParseSplittedCoordinates(coordinatesValue.Replace(" ", string.Empty).Split(','));
        }

        /// <summary>
        /// Inserts <see cref="ConfigExpression"/> instance to expression collection.
        /// </summary>
        /// <param name="expression"><see cref="ConfigExpression"/> instance that describes single search expression.</param>
        /// <exception cref="ArgumentNullException">Empty config expression.</exception>
        public void AddSearchExpression(ConfigExpression expression)
        {
            if (expression != null)
            {
                this.Expressions.Add(expression);
            }
            else
            {
                throw new ArgumentNullException($"Null config expression for field {this.Name} was provided.");
            }
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"Config field: {this.Name}; value type: {this.ValueType}";
        }

        private static Tuple<int, int> ParseSplittedCoordinates(string[] coordinates)
        {
            return new Tuple<int, int>(
                TryParseStringAsInteger(coordinates[0]),
                TryParseStringAsInteger(coordinates[1]));
        }

        private static int TryParseStringAsInteger(string value)
        {
            return int.TryParse(value, out int result)
                ? result
                : -1;
        }

        private static void JoinSplittedRegularExpression(List<string> splittedValue)
        {
            while (splittedValue.Count != 2)
            {
                MergeSplittedPattern(splittedValue);
            }
        }

        private static void MergeSplittedPattern(List<string> splittedValue)
        {
            splittedValue[0] = $"{splittedValue[0]};{splittedValue[1]}";
            OffsetInputByOneValueToLeft(splittedValue);
            splittedValue.RemoveAt(splittedValue.Count - 1);
        }

        private static void OffsetInputByOneValueToLeft(List<string> splittedValue)
        {
            for (int i = 2; i < splittedValue.Count; i++)
            {
                splittedValue[i - 1] = splittedValue[i];
            }
        }

        private static string EncodeString(string value)
        {
            return new DefaultSoundexEncoder(value).EncodedValue;
        }

        private void SplitInputAndSetProperties(string input)
        {
            var splittedValue = input.Split(';').ToList();
            JoinSplittedRegularExpression(splittedValue);
            this.SetPropertiesFromSplittedValue(splittedValue);
        }

        private void SetPropertiesFromSplittedValue(List<string> splittedValue)
        {
            this.TextExpression = this.ValidateValueOrGetName(splittedValue[0]);
            this.ExpectedName = this.ValidateValueOrGetName(splittedValue[1]);
            this.ValidateSoundexStatus();
        }

        private string ValidateValueOrGetName(string value)
        {
            return string.IsNullOrEmpty(value)
                ? this.Name
                : value;
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
            this.ExpectedName = EncodeString(this.ExpectedName);
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
