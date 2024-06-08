using ExcelMvc.Functions;
using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class FunctionDiscoveryTests
    {
        public FunctionDiscoveryTests()
        {
            FunctionHost.Instance = new ExcelFunctionHost();
        }

        [Function(Name = "uAdd3",
            Description = nameof(uAdd3),
            HelpTopic ="https://microsoft.com",
            IsAsync = true,
            IsThreadSafe = true,
            IsVolatile = true,
            IsClusterSafe = true,
            IsHidden = true,
            IsMacroType = true)]
        public static double uAdd3(double v1, double v2, double v3)
        {
            return v1 + v2 + v3;
        }

        [Function(Name = "uAdd2",
            Description = nameof(uAdd3),
            HelpTopic = "https://microsoft.com",
            IsAsync = false,
            IsThreadSafe = false,
            IsVolatile = false,
            IsClusterSafe = false,
            IsHidden = false,
            IsMacroType = false)]
        public static double uAdd2(double v1, double v2)
        {
            return v1 + v2;
        }

        [TestMethod]
        public void Discover()
        {
            var functions = FunctionDiscovery.DiscoverFunctions().ToArray();
            Assert.AreEqual(2, functions.Length);
        }
    }
}