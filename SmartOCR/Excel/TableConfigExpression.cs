using System.Collections.Generic;
using System.Globalization;

namespace SmartOCR
{
    public class TableConfigExpression : ConfigExpressionBase
    {
        public int Row { get; set; }
        public int Column { get; set; }

        public TableConfigExpression(string input)
        {
            ParseInput(input);
        }
        protected override List<string> ParseInput(string input)
        {
            List<string> splitted_input = base.ParseInput(input);
            Row = string.IsNullOrEmpty(splitted_input[1])
                ? 0
                : int.Parse(splitted_input[1], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
            Column = string.IsNullOrEmpty(splitted_input[2])
                ? 0
                : int.Parse(splitted_input[2], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
            return splitted_input;
        }
    }
}
