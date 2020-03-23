using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SmartOCR.Soundex;

namespace SmartOCRTesting
{
    [TestClass]
    public class DaitchMokotoffSoundexTest
    {
        string Encode(string value)
        {
            return new DaitchMokotoffSoundexEncoder(value).EncodedValue;
        }

        [TestMethod]
        public void SingleWordTest()
        {
            Assert.AreEqual("463000", Encode("Schmidt"));
        }

        [TestMethod]
        public void SentenceTest()
        {
            Assert.AreEqual("300000 550000 797600 754000", Encode("The quick brown fox"));
        }

        [TestMethod]
        public void CyrillicAndLatinWordTest()
        {
            string first = Encode("Арнольд Шварцнеггер");
            string second = Encode("Arnold Schwarzenegger");
            Assert.AreEqual(first, second);
        }
    }
}
