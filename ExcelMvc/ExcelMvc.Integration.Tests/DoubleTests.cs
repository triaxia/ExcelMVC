using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class DoubleTests
    {
        [Function()]
        public static double uDouble(double v1, [Argument(Name = "[v2]")] double? v2 = 0)
        {
            return (double)(v1 - v2.Value);
        }

        [TestMethod]
        public void uDouble()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (double)excel.Application.Run("uDouble", double.MaxValue);
                Assert.AreEqual(double.MaxValue, result);
                result = (double)excel.Application.Run("uDouble", 123.3456, 123.000);
                Assert.AreEqual(123.3456 - 123.000, result);
            }
        }
    }
}