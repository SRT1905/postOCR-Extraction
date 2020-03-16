namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Encodes string using Daitch-Mokotoff soundex algorithm.
    /// </summary>
    public class DaitchMokotoffSoundexEncoder : Soundex
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DaitchMokotoffSoundexEncoder"/> class.
        /// </summary>
        /// <param name="value">A string to encode.</param>
        public DaitchMokotoffSoundexEncoder(string value)
        {
            this.SourceValue = value;
            this.EncodedValue = this.EncodeValue(value);
        }

        /// <inheritdoc/>
        protected override string EncodeSingleWord(string word) // TODO: add Daitch-Mokotoff Soundex implementation
        {
            throw new NotImplementedException();
        }
    }
}
