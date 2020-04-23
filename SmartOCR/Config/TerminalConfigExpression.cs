namespace SmartOCR.Config
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Describes custom terminal search expression.
    /// </summary>
    public class TerminalConfigExpression : ConfigExpression
    {
        private const string SpecialParameterPattern = @"special\((.*)\)";

        /// <summary>
        /// Initializes a new instance of the <see cref="TerminalConfigExpression"/> class.
        /// Instance is initialized by Excel cell contents.
        /// </summary>
        /// <param name="input">Excel cell contents, containing regular expression pattern, line offset and horizontal search status.</param>
        /// <param name="valueType">String representation of field data type.</param>
        public TerminalConfigExpression(string valueType, string input)
            : base(valueType, input)
        {
        }

        /// <summary>
        /// Gets or sets parameter for custom processing of terminal node.
        /// </summary>
        public string TerminalParameter { get; set; } = string.Empty;

        /// <inheritdoc/>
        protected override List<string> ParseInput(string input)
        {
            if (Regex.IsMatch(input, SpecialParameterPattern))
            {
                input = this.ExtractTerminalParameter(input);
            }

            return base.ParseInput(input);
        }

        private string ExtractTerminalParameter(string input)
        {
            this.TerminalParameter = Regex.Match(input, SpecialParameterPattern).Captures[1].Value;
            return Regex.Replace(input, SpecialParameterPattern, string.Empty);
        }
    }
}
