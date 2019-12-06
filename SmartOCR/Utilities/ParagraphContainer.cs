using Microsoft.Office.Interop.Word;
using System;

namespace SmartOCR
{
    internal class ParagraphContainer : IComparable<ParagraphContainer>
    {
        public double HorizontalLocation;
        public double VerticalLocation;
        public string Text;

        public ParagraphContainer(Range range)
        {
            HorizontalLocation = ValidateLocation(range, WdInformation.wdHorizontalPositionRelativeToPage);
            VerticalLocation = ValidateLocation(range, WdInformation.wdVerticalPositionRelativeToPage);
            Text = RemoveInvalidChars(range.Text);
        }

        private double ValidateLocation(Range range, WdInformation information)
        {
            return range.Information[information] != -1
                ? range.Information[information]
                : range.Words[1].Information[information];
        }

        private string RemoveInvalidChars(string check_string)
        {
            string[] separators = new string[] { "\r", "\a", "\t", "\f" };
            string[] temp = check_string.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return string.Join("", temp);
        }

        public int CompareTo(ParagraphContainer that)
        {
            if (that == null)
            {
                return 1;
            }
            return HorizontalLocation.CompareTo(that.HorizontalLocation);
        }
    }
}