using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class ValueTests
    {
        [Function()]
        public static bool uValueMissing([Argument(Name = "[v1]")] object v1)
        {
            return v1 == FunctionHost.Instance.ValueMissing;
        }

        [TestMethod]
        public void uValueMissing()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (bool)excel.Application.Run("uValueMissing");
                Assert.AreEqual(true, result);
                 result = (bool)excel.Application.Run("uValueMissing", 123);
                Assert.AreEqual(false, result);
            }
        }

        [Function()]
        public static bool uValueEmpty([Argument(Name = "[v1]")] object v1)
        {
            return $"{v1}" == "";
        }

        [TestMethod]
        public void uValueEmpty()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (bool)excel.Application.Run("uValueEmpty", FunctionHost.Instance.ValueEmpty);
                Assert.AreEqual(true, result);
            }
        }

        [Function()]
        public static object uGetErrorNA()
        {
            return FunctionHost.Instance.ErrorNA;
        }

        [Function()]
        public static bool uSetErrorNA([Argument(Name = "[v1]")] object v1)
        {
            //Debugger.Launch();
            return v1.Equals(FunctionHost.Instance.ErrorNA) || ((int)(double) v1) == -2146826246;
        }

        [TestMethod]
        public void uErrorNA()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (object)excel.Application.Run("uGetErrorNA");
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorNA", result));
            }
        }
    }
}
