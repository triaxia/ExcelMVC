using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class NullTests
    {
        [Function()]
        public static object[] uEmptyArray([Argument(Name = "[v1]")] object v1 = null)
        {
            return new object[] { };
        }

        [TestMethod]
        public void uEmptyArray()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (object)excel.Application.Run("uEmptyArray");
                Assert.AreEqual(FunctionHost.Instance.ErrorNumbers[FunctionHost.Instance.ErrorNum], result);
            }
        }

        [Function()]
        public static object[] uNullArray([Argument(Name = "[v1]")] object v1 = null)
        {
            return null;
        }
        
        [TestMethod]
        public void uNullArray()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (object)excel.Application.Run("uNullArray");
                Assert.AreEqual(FunctionHost.Instance.ErrorNumbers[FunctionHost.Instance.ErrorNum], result);
            }
        }

        [Function()]
        public static object uNullValue([Argument(Name = "[v1]")] object v1 = null)
        {
            return null;
        }

        [TestMethod]
        public void uNullValue()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (object)excel.Application.Run("uNullValue");
                Assert.AreEqual(FunctionHost.Instance.ErrorNumbers[FunctionHost.Instance.ErrorNum], result);
            }
        }

    }
}