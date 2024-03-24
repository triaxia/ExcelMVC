using ExcelMvc.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class DelegateFactoryTests
    {
        public static double Add(double x, double y)
        {
            return x + y;
        }

        [TestMethod]
        public void Discover()
        {
            var method = typeof(DelegateFactoryTests).GetMethod("Add");
            var dele = DelegateFactory.MakeInnerDelegate(method);
        }
    }
}
