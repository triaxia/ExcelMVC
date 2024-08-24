using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class ShortTests
    {
        [Function()]
        public static ushort uUShort(ushort v1, [Argument(Name = "[v2]")] ushort? v2 = 0)
        {
            return (ushort)(v1 + v2.Value);
        }

        [TestMethod]
        public void uUShort()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (ushort)excel.Application.Run("uUShort", ushort.MaxValue);
                Assert.AreEqual(ushort.MaxValue, result);
                var half = ushort.MaxValue / 2;
                result = (ushort)excel.Application.Run("uUShort", half, half);
                Assert.AreEqual(half * 2, result);
            }
        }

        [Function()]
        public static short uShort(short v1, [Argument(Name = "[v2]")] short? v2 = 0)
        {
            return (short)(v1 - v2.Value);
        }

        [TestMethod]
        public void uShort()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (short)excel.Application.Run("uShort", short.MaxValue);
                Assert.AreEqual(short.MaxValue, result);
                var half = short.MaxValue / 2;
                result = (short)excel.Application.Run("uShort", half - 1, half);
                Assert.AreEqual(-1, result);
            }
        }
    }
}