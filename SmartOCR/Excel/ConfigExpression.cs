using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartOCR
{
    class ConfigExpression
    {
        public string RE_Pattern { get; set; }
        public int LineOffset { get; set; }
        public int HorizontalStatus { get; set; }
        // 0 - search whole line;
        // 1 - search only paragraphs to the right
        // -1 - search only paragrapgs to the left

        public ConfigExpression(string pattern, int line_offset, int position_status)
        {
            RE_Pattern = pattern;
            LineOffset = line_offset;
            HorizontalStatus = -1 <= position_status && position_status <= 1
                ? position_status
                : throw new ArgumentOutOfRangeException("Horizontal position status must be in range [-1;1].");
        }

        public ConfigExpression(string input)
        {
            ParseInput(input);
        }

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
        }
    }
}
