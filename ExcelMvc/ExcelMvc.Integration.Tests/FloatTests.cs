using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class FloatTests
    {
        [Function()]
        public static double uFloat(float v1, [Argument(Name = "[v2]")] float? v2 = 0)
        {
            return (double)(v1 - v2.Value);
        }

        [TestMethod]
        public void uFloat()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (float)excel.Application.Run("uFloat", float.MaxValue);
                Assert.AreEqual(float.MaxValue, result);
                float f1 = 123.3456F;
                float f2 = 123F;
                result = (float)excel.Application.Run("uFloat", f1, f2);
                Assert.AreEqual(f1 - f2, result);
            }
        }
    }
}