using ExcelMvc.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class FunctionDiscoveryTests
    {
        [ExcelFunction(Name = "uAdd3",
            Description = nameof(uAdd3),
            HelpTopic ="https://microsoft.com",
            IsAsync = true,
            IsThreadSafe = true,
            IsVolatile = true,
            IsClusterSafe = true,
            IsHidden = true,
            IsMacro = true)]
        public static double uAdd3(double v1, double v2, double v3)
        {
            return v1 + v2 + v3;
        }

        [ExcelFunction(Name = "uAdd2",
            Description = nameof(uAdd3),
            HelpTopic = "https://microsoft.com",
            IsAsync = false,
            IsThreadSafe = false,
            IsVolatile = false,
            IsClusterSafe = false,
            IsHidden = false,
            IsMacro = false)]
        public static double uAdd2(double v1, double v2)
        {
            return v1 + v2;
        }

        [TestMethod]
        public void Discover()
        {
            var functions = FunctionDiscovery.Discover().ToArray();
            Assert.AreEqual(2, functions.Length);
        }
    }
}