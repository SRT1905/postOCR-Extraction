using System;
using Microsoft.Office.Interop.Word;

namespace SmartOCR
{
    class ParagraphContainer
    {
        public double HorizontalLocation;
        public double VerticalLocation;
        public string Text;

        public ParagraphContainer(Range range)
        {
            HorizontalLocation = range.Information[WdInformation.wdHorizontalPositionRelativeToPage];
            if (HorizontalLocation == -1)
            {
                HorizontalLocation = range.Words[1].Information[WdInformation.wdHorizontalPositionRelativeToPage];
            }

            VerticalLocation = range.Information[WdInformation.wdVerticalPositionRelativeToPage];
            if (VerticalLocation == -1)
            {
                VerticalLocation = range.Words[1].Information[WdInformation.wdVerticalPositionRelativeToPage];
            }

            Text = RemoveInvalidChars(range.Text);
        }
        private string RemoveInvalidChars(string check_string)
        {
            string[] separators = new string[] {"\r", "\a", "\t", "\f"};
            string[] temp = check_string.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return string.Join("", temp);
        }
    }
}