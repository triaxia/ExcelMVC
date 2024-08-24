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
                var result = (byte)excel.Application.Run("uByte", 123);
                Assert.AreEqual(123, result);
                result = (byte)excel.Application.Run("uByte", 123, 123);
                Assert.AreEqual(246, result);
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
                var result = (sbyte)excel.Application.Run("uSByte", 123);
                Assert.AreEqual(123, result);
                result = (sbyte)excel.Application.Run("uSByte", 120, 123);
                Assert.AreEqual(-3, result);
            }
        }
    }
}