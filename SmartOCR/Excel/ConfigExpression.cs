using System;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// Describes single search expression defined in Excel config file.
    /// </summary>
    class ConfigExpression
    {
        /// <summary>
        /// Describes regular expression pattern.
        /// </summary>
        public string RE_Pattern { get; set; }

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

        /// <summary>
        /// Initializes a new <see cref="ConfigExpression"/> instance with Excel cell contents.
        /// </summary>
        /// <param name="input">Excel cell contents, containing regular expression pattern, line offset and horizontal search status.</param>
        public ConfigExpression(string input)
        {
            ParseInput(input);
        }

        /// <summary>
        /// Extracts data from Excel cell contents.
        /// </summary>
        /// <param name="input">Excel cell contents, containing regular expression pattern, line offset and horizontal search status.</param>
        private void ParseInput(string input)
        {
            var splitted_input = input.Split(';').ToList();
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
                splitted_input.Add(null);
            }

            RE_Pattern = splitted_input[0];
            LineOffset = string.IsNullOrEmpty(splitted_input[1])
                ? 0
                : int.Parse(splitted_input[1]);
            HorizontalStatus = string.IsNullOrEmpty(splitted_input[2])
                ? 0
                : int.Parse(splitted_input[2]);
            
            if (Math.Abs(HorizontalStatus) > 1)
            {
                throw new ArgumentOutOfRangeException("Horizontal position status must be in range [-1; 1].");
            }
        }
    }
}
