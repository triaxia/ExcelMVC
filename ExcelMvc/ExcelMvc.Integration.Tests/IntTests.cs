using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class IntTests
    {
        [Function()]
        public static uint uUInt(uint v1, [Argument(Name = "[v2]")] uint? v2 = 0)
        {
            return (uint)(v1 + v2.Value);
        }

        [TestMethod]
        public void uUInt()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (uint)excel.Application.Run("uUInt", uint.MaxValue);
                Assert.AreEqual(uint.MaxValue, result);
                var half = uint.MaxValue / 2;
                result = (uint)excel.Application.Run("uUInt", half, half);
                Assert.AreEqual(half * 2, result);
            }
        }

        [Function()]
        public static int uInt(int v1, [Argument(Name = "[v2]")] int? v2 = 0)
        {
            return (int)(v1 - v2.Value);
        }

        [TestMethod]
        public void uInt()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (int)excel.Application.Run("uInt", int.MaxValue);
                Assert.AreEqual(int.MaxValue, result);
                var half = int.MaxValue / 2;
                result = (int)excel.Application.Run("uInt", half - 1, half);
                Assert.AreEqual(-1, result);
            }
        }
    }
}