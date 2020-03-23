using Microsoft.VisualStudio.TestTools.UnitTesting;
using SmartOCR.Soundex;

namespace SmartOCRTesting
{
    [TestClass]
    public class DefaultSoundexEncoderTest
    {

        string Encode(string value)
        {
            return new DefaultSoundexEncoder(value).EncodedValue;
        }

        [TestMethod]
        public void SingleWordTest()
        {
            Assert.AreEqual(Encode("Invoice"), "I512");
        }

        [TestMethod]
        public void SentenceSplitBySpacesTest()
        {
            Assert.AreEqual(Encode("Total amount"), "T340 A553");
        }

        [TestMethod]
        public void SentenceSplitByPunctuationAndSpacesTest()
        {
            Assert.AreEqual(
                Encode("Since soundex is based on English pronunciation, some European names may not soundex correctly."),
                "S520 S532 I200 B230 O500 E524 P655, S500 E615 N520 M000 N300 S532 C623.");
        }
    }
}
