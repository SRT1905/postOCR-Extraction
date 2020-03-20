namespace SmartOCRTesting
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using SmartOCR.Search.SimilarityAlgorithms;

    [TestClass]
    public class JaroWrinklerAlgorithmTest
    {
        static readonly JaroWrinklerAlgorithm algorithm = new JaroWrinklerAlgorithm();

        [TestMethod]
        public void EqualStringTest()
        {
            Assert.AreEqual(1, algorithm.GetStringSimilarity("INVOICE", "INVOICE"));
        }

        [TestMethod]
        public void CompletelyDifferentStringTest()
        {
            Assert.IsTrue(algorithm.GetStringSimilarity("undefined", "he-who-must-not-be-named") >= 0.4);
        }

        [TestMethod]
        public void AlmostSimilarStringsTest()
        {
            Assert.IsTrue(algorithm.GetStringSimilarity("Грузоотправитель", "Грузотравиель") >= 0.75);
        }
    }
}
