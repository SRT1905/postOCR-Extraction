namespace SmartOCR
{
    using System;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Container class that is used to store data from single Word document paragraph.
    /// </summary>
    public class ParagraphContainer : IComparable<ParagraphContainer>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ParagraphContainer"/> class.
        /// Instance stores information from <see cref="Range"/> object.
        /// </summary>
        /// <param name="range">Representation of single Word document paragraph.</param>
        public ParagraphContainer(Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            this.HorizontalLocation = ValidateLocation(range, WdInformation.wdHorizontalPositionRelativeToPage);
            this.VerticalLocation = ValidateLocation(range, WdInformation.wdVerticalPositionRelativeToPage);
            this.Text = RemoveInvalidChars(range.Text);
            this.Soundex = new DefaultSoundexEncoder(this.Text).EncodedValue;
        }

        /// <summary>
        /// Gets distance from left edge of paragraph to left edge of page.
        /// Measured in points (72 point = 1 inch).
        /// </summary>
        public decimal HorizontalLocation { get; }

        /// <summary>
        /// Gets Soundex equivalent of <see cref="ParagraphContainer.Text"/>.
        /// </summary>
        public string Soundex { get; }

        /// <summary>
        /// Gets paragraph text.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Gets distance from top edge of paragraph to top edge of page.
        /// Measured in points (72 point = 1 inch).
        /// </summary>
        public decimal VerticalLocation { get; }

        public static bool operator ==(ParagraphContainer left, ParagraphContainer right)
        {
            return left?.Equals(right) ?? right is null;
        }

        public static bool operator !=(ParagraphContainer left, ParagraphContainer right)
        {
            return !(left == right);
        }

        public static bool operator <(ParagraphContainer left, ParagraphContainer right)
        {
            return left is null ? !(right is null) : left.CompareTo(right) < 0;
        }

        public static bool operator <=(ParagraphContainer left, ParagraphContainer right)
        {
            return left is null || left.CompareTo(right) <= 0;
        }

        public static bool operator >(ParagraphContainer left, ParagraphContainer right)
        {
            return !(left is null) && left.CompareTo(right) > 0;
        }

        public static bool operator >=(ParagraphContainer left, ParagraphContainer right)
        {
            return left is null ? right is null : left.CompareTo(right) >= 0;
        }

        /// <inheritdoc/>
        public int CompareTo(ParagraphContainer that)
        {
            return that == null
                ? 1
                : this.HorizontalLocation.CompareTo(that.HorizontalLocation);
        }

        /// <inheritdoc/>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        /// <inheritdoc/>
        public override bool Equals(object obj)
        {
            return ReferenceEquals(this, obj) || (!(obj is null) && this.CompareTo((ParagraphContainer)obj) == 0);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"X: {this.HorizontalLocation}; Y: {this.VerticalLocation}; Text: {this.Text}";
        }

        /// <summary>
        /// Removes characters that interfere with meaningful text content.
        /// </summary>
        /// <param name="checkString">Text to process.</param>
        /// <returns>Paragraph text, cleansed of invalid characters.</returns>
        private static string RemoveInvalidChars(string checkString)
        {
            string[] temp = checkString.Split(new[] { "\r", "\a", "\t", "\f" }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(string.Empty, temp).Replace("\v", " ");
        }

        /// <summary>
        /// Gets paragraph location specified by <see cref="WdInformation"/> enumeration.
        /// </summary>
        /// <param name="range">Representation of single Word document paragraph.</param>
        /// <param name="information"><see cref="WdInformation"/> object that represents type of returned location.</param>
        /// <returns>Position of paragraph within document page.</returns>
        private static decimal ValidateLocation(Range range, WdInformation information)
        {
            decimal temp = (decimal)range.Words[1].Information[information];
            return decimal.Round(temp, 1);
        }
    }
}