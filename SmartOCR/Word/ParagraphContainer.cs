using Microsoft.Office.Interop.Word;
using System;

namespace SmartOCR
{
    internal class ParagraphContainer : IComparable<ParagraphContainer>
    {
        public decimal HorizontalLocation;
        public decimal VerticalLocation;
        public string Text;

        public ParagraphContainer(Range range)
        {
            HorizontalLocation = ValidateLocation(range, WdInformation.wdHorizontalPositionRelativeToPage);
            VerticalLocation = ValidateLocation(range, WdInformation.wdVerticalPositionRelativeToPage);
            Text = RemoveInvalidChars(range.Text);
        }

        private decimal ValidateLocation(Range range, WdInformation information)
        {
            decimal temp = (decimal)range.Words[1].Information[information];
            return decimal.Round(temp, 1);
        }

        private string RemoveInvalidChars(string check_string)
        {
            string[] separators = new string[] { "\r", "\a", "\t", "\f"};
            string[] temp = check_string.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return string.Join("", temp).Replace("\v", " ");
        }

        public int CompareTo(ParagraphContainer that)
        {
            if (that == null)
            {
                return 1;
            }
            return HorizontalLocation.CompareTo(that.HorizontalLocation);
        }

        public override string ToString()
        {
            return $"X: {HorizontalLocation}; Y: {VerticalLocation}; Text: {Text}";
        }
    }
}