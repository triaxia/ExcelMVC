using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class BoolTests
    {
        [Function()]
        public static bool uBool(bool v1, [Argument(Name = "[v2]")] bool? v2 = false)
        {
            return v1 && v2.Value;
        }

        [TestMethod]
        public void uBool()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (bool)excel.Application.Run("uBool", true);
                Assert.IsFalse(result);
                result = (bool)excel.Application.Run("uBool", true, true);
                Assert.IsTrue(result);
            }
        }
    }
}