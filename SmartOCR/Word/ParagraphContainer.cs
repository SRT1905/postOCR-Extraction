using Microsoft.Office.Interop.Word;
using System;

namespace SmartOCR
{
    /// <summary>
    /// Container class that is used to store data from single Word document paragraph.
    /// </summary>
    public class ParagraphContainer : IComparable<ParagraphContainer>
    {
        /// <summary>
        /// Indicates distance from left edge of paragraph to left edge of page.
        /// Measured in points (72 point = 1 inch).
        /// </summary>
        public decimal HorizontalLocation { get; }

        /// <summary>
        /// Indicates distance from top edge of paragraph to top edge of page.
        /// Measured in points (72 point = 1 inch).
        /// </summary>
        public decimal VerticalLocation { get; }

        /// <summary>
        /// Paragraph text.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Initializes instance of <see cref="ParagraphContainer"/> object that stores information from <see cref="Range"/> object.
        /// </summary>
        /// <param name="range">Representation of single Word document paragraph.</param>
        public ParagraphContainer(Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }
            HorizontalLocation = ValidateLocation(range, WdInformation.wdHorizontalPositionRelativeToPage);
            VerticalLocation = ValidateLocation(range, WdInformation.wdVerticalPositionRelativeToPage);
            Text = RemoveInvalidChars(range.Text);
        }

        /// <summary>
        /// Initializes instance of <see cref="ParagraphContainer"/> object that stores provided locations and text.
        /// </summary>
        /// <param name="range">Representation of single Word document paragraph.</param>
        public ParagraphContainer(double horizontalLocation, double verticalLocation, string text)
        {
            HorizontalLocation = ValidateLocation(horizontalLocation);
            VerticalLocation = ValidateLocation(verticalLocation);
            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }
            Text = RemoveInvalidChars(text);
        }

        /// <summary>
        /// Gets paragraph location specified by <see cref="WdInformation"/> enumeration.
        /// </summary>
        /// <param name="range">Representation of single Word document paragraph.</param>
        /// <param name="information"><see cref="WdInformation"/> object that represents type of returned location.</param>
        /// <returns>Position of paragraph within document page.</returns>
        private decimal ValidateLocation(Range range, WdInformation information)
        {
            decimal temp = (decimal)range.Words[1].Information[information];
            return decimal.Round(temp, 1);
        }

        /// <summary>
        /// Casts provided number to <see cref="decimal"/>.
        /// </summary>
        /// <param name="number">Number that represents paragraph location on page.</param>
        /// <returns>Position of paragraph within document page.</returns>
        private decimal ValidateLocation(double number)
        {
            decimal temp = (decimal)number;
            return decimal.Round(temp, 1);
        }

        /// <summary>
        /// Removes characters that interfere with meaningful text content.
        /// </summary>
        /// <param name="check_string">Text to process.</param>
        /// <returns>Paragraph text, cleansed of invalid characters.</returns>
        private string RemoveInvalidChars(string check_string)
        {
            string[] separators = new string[] { "\r", "\a", "\t", "\f" };
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

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
            {
                return true;
            }

            if (obj is null)
            {
                return false;
            }

            return this.CompareTo((ParagraphContainer)obj) == 0;
        }

        public static bool operator ==(ParagraphContainer left, ParagraphContainer right)
        {
            return left is null ? right is null : left.Equals(right);
        }

        public static bool operator !=(ParagraphContainer left, ParagraphContainer right)
        {
            return !(left == right);
        }

        public static bool operator <(ParagraphContainer left, ParagraphContainer right)
        {
            return left is null ? right is object : left.CompareTo(right) < 0;
        }

        public static bool operator <=(ParagraphContainer left, ParagraphContainer right)
        {
            return left is null || left.CompareTo(right) <= 0;
        }

        public static bool operator >(ParagraphContainer left, ParagraphContainer right)
        {
            return left is object && left.CompareTo(right) > 0;
        }

        public static bool operator >=(ParagraphContainer left, ParagraphContainer right)
        {
            return left is null ? right is null : left.CompareTo(right) >= 0;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}