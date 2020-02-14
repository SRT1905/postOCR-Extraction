using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    class ConfigExpressionBase
    {
        /// <summary>
        /// Describes regular expression pattern.
        /// </summary>
        public string RE_Pattern { get; set; }

        protected virtual List<string> ParseInput(string input)
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
            return splitted_input;
        }
    }
}
