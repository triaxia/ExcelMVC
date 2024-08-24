using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class ShortTests
    {
        [Function()]
        public static ushort uShort(ushort v1, [Argument(Name = "[v2]")] ushort? v2 = 0)
        {
            return (ushort)(v1 + v2.Value);
        }

        [TestMethod]
        public void uShort()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (ushort)excel.Application.Run("uShort", ushort.MaxValue);
                Assert.AreEqual(ushort.MaxValue, result);
                var half = ushort.MaxValue / 2;
                result = (ushort)excel.Application.Run("uShort", half, half);
                Assert.AreEqual(half * 2, result);
            }
        }

        [Function()]
        public static short uSShort(short v1, [Argument(Name = "[v2]")] short? v2 = 0)
        {
            return (short)(v1 - v2.Value);
        }

        [TestMethod]
        public void uSShort()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (short)excel.Application.Run("uSShort", short.MaxValue);
                Assert.AreEqual(short.MaxValue, result);
                var half = short.MaxValue / 2;
                result = (short)excel.Application.Run("uSShort", half - 1, half);
                Assert.AreEqual(-1, result);
            }
        }
    }
}