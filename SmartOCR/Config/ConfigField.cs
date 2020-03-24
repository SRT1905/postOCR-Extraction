namespace SmartOCR.Config
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using SmartOCR.Soundex;
    using SmartOCR.Word;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

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

            this.GridCoordinates = ParseSplitCoordinates(coordinatesValue.Replace(" ", string.Empty).Split(','));
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
        public override string ToString() => string.Format("Config field: {0}; value type: {1}", this.Name, this.ValueType);

        private static Tuple<int, int> ParseSplitCoordinates(string[] coordinates)
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

        private static void JoinSplitRegularExpression(List<string> splitValue)
        {
            while (splitValue.Count != 2)
            {
                MergeSplitPattern(splitValue);
            }
        }

        private static void MergeSplitPattern(List<string> splitValue)
        {
            splitValue[0] = $"{splitValue[0]};{splitValue[1]}";
            OffsetInputByOneValueToLeft(splitValue);
            splitValue.RemoveAt(splitValue.Count - 1);
        }

        private static void OffsetInputByOneValueToLeft(List<string> splitValue)
        {
            for (int i = 2; i < splitValue.Count; i++)
            {
                splitValue[i - 1] = splitValue[i];
            }
        }

        private static string EncodeString(string value) => new DaitchMokotoffSoundexEncoder(value).EncodedValue;

        private void SplitInputAndSetProperties(string input)
        {
            var splitValue = input.Split(';').ToList();
            JoinSplitRegularExpression(splitValue);
            this.SetPropertiesFromSplitValue(splitValue);
        }

        private void SetPropertiesFromSplitValue(List<string> splitValue)
        {
            this.TextExpression = this.ValidateValueOrGetName(splitValue[0]);
            this.ExpectedName = this.ValidateValueOrGetName(splitValue[1]);
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

        private string ExtractSoundexExpression() => Utilities.CreateRegexpObject(@"soundex\((.*)\)")
                                                              .Match(this.TextExpression)
                                                              .Groups[1]
                                                              .Value;
    }
}
