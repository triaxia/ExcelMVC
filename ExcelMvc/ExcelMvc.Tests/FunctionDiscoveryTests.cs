using ExcelMvc.Functions;
using Mvc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class FunctionDiscoveryTests
    {
        [Function(Name = "uAdd")]
        public static double uAdd(double v1, double v2, double v3)
        {
            return v1 + v2 + v3;
        }

        [TestMethod]
        public void Discover()
        {
            var functions = FunctionDiscovery.Discover().Where(x => x.function.Name == "uAdd");
            Assert.AreEqual(1, functions.Count());
        }
    }
}