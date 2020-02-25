using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartOCR
{
    /// <summary>
    /// Describes single search expression defined in Excel config file.
    /// </summary>
    public class ConfigExpression : ConfigExpressionBase
    {
        #region Properties
        /// <summary>
        /// Shows how offset should be search line, comparing to one, where previous search pattern was matched.
        /// </summary>
        public int LineOffset { get; set; }
        /// <summary>
        /// Indicates how single line contents should be searched. 
        /// <para>0 - search whole line.</para>
        /// <para>1 - search only paragraphs to the right of previously found match.</para>
        /// <para>-1 - search only paragrapgs to the left of previously found match.</para>
        /// </summary>
        public int HorizontalStatus { get; set; }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new <see cref="ConfigExpression"/> instance with Excel cell contents.
        /// </summary>
        /// <param name="input">Excel cell contents, containing regular expression pattern, line offset and horizontal search status.</param>
        public ConfigExpression(string input)
        {
            ParseInput(input);
        }
        #endregion

        #region Protected methods
        /// <summary>
        /// Extracts data from Excel cell contents.
        /// </summary>
        /// <param name="input">Excel cell contents, containing regular expression pattern, line offset and horizontal search status.</param>
        protected override List<string> ParseInput(string input)
        {
            List<string> splitted_input = base.ParseInput(input);
            LineOffset = string.IsNullOrEmpty(splitted_input[1])
                ? 0
                : int.Parse(splitted_input[1], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
            HorizontalStatus = string.IsNullOrEmpty(splitted_input[2])
                ? 0
                : int.Parse(splitted_input[2], NumberStyles.Any, NumberFormatInfo.CurrentInfo);

            if (Math.Abs(HorizontalStatus) > 1)
            {
                throw new ArgumentOutOfRangeException(nameof(input), Properties.Resources.outOfRangeParagraphHorizontalLocationStatus);
            }
            return splitted_input;
        }
        #endregion
    }
}
