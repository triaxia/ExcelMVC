using ExcelMvc.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [ExcelFunction(Name = "uAdd")]
        public static double Add(double v1, double v2, double v3)
        {
            return v1 + v2 + v3;
        }

        [TestMethod]
        public void TestMethod1()
        {
            var x = FunctionDiscovery.Discover().First();
            for (int idx = 0; idx < 1000000; idx++)
                XlCall.Register(x.function);
        }
    }
}