using System.Collections.Generic;
using System.Globalization;

namespace SmartOCR
{
    class TableConfigExpression : ConfigExpressionBase
    {
        public int RowOffset { get; set; }
        public int ColumnOffset { get; set; }

        public TableConfigExpression(string input)
        {
            ParseInput(input);
        }
        protected override List<string> ParseInput(string input)
        {
            List<string> splitted_input = base.ParseInput(input);
            RowOffset = string.IsNullOrEmpty(splitted_input[1])
                ? 0
                : int.Parse(splitted_input[1], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
            ColumnOffset = string.IsNullOrEmpty(splitted_input[2])
                ? 0
                : int.Parse(splitted_input[2], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
            return splitted_input;
        }
    }
}
