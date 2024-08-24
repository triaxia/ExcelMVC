using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class ByteTests
    {
        [Function()]
        public static byte uByte(byte v1, [Argument(Name = "[v2]")] byte? v2 = 0)
        {
            return (byte)(v1 + v2.Value);
        }

        [TestMethod]
        public void uByte()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (byte)excel.Application.Run("uByte", byte.MaxValue);
                Assert.AreEqual(byte.MaxValue, result);
                var half = byte.MaxValue / 2;
                result = (byte)excel.Application.Run("uByte", half, half);
                Assert.AreEqual(half * 2, result);
            }
        }

        [Function()]
        public static sbyte uSByte(sbyte v1, [Argument(Name = "[v2]")] sbyte? v2 = 0)
        {
            return (sbyte)(v1 - v2.Value);
        }

        [TestMethod]
        public void uSByte()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (sbyte)excel.Application.Run("uSByte", sbyte.MaxValue);
                Assert.AreEqual(sbyte.MaxValue, result);
                var half = byte.MaxValue / 2;
                result = (sbyte)excel.Application.Run("uSByte", half - 1, half);
                Assert.AreEqual(-1, result);
            }
        }
    }
}