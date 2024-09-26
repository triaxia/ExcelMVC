using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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

        private static bool CompareError(object arg, object error)
        {
            return arg.Equals(error) // from Excel
                || (FunctionHost.Instance.ErrorNumbers.TryGetValue(error, out var num) && ((int)(double)arg) == num); // from C#
        }

        [Function()]
        public static object uGetErrorNA()
            => FunctionHost.Instance.ErrorNA;

        [Function()]
        public static bool uSetErrorNA([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorNA); 

        [TestMethod]
        public void uErrorNA()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorNA"
                    , (object)excel.Application.Run("uGetErrorNA")));
        }

        [Function()]
        public static object uGetErrorNull()
            => FunctionHost.Instance.ErrorNull;

        [Function()]
        public static bool uSetErrorNull([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorNull);

        [TestMethod]
        public void uErrorNull()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorNull",
                    (object)excel.Application.Run("uGetErrorNull")));
        }

        [Function()]
        public static object uGetErrorName()
            => FunctionHost.Instance.ErrorName;

        [Function()]
        public static bool uSetErrorName([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorName);

        [TestMethod]
        public void uErrorName()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorName",
                    (object)excel.Application.Run("uGetErrorName")));
        }

        [Function()]
        public static object uGetErrorDiv0()
            => FunctionHost.Instance.ErrorDiv0;

        [Function()]
        public static bool uSetErrorDiv0([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorDiv0);

        [TestMethod]
        public void uErrorDiv0()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorDiv0",
                    (object)excel.Application.Run("uGetErrorDiv0")));
        }

        [Function()]
        public static object uGetErrorRef()
            => FunctionHost.Instance.ErrorRef;

        [Function()]
        public static bool uSetErrorRef([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorRef);

        [TestMethod]
        public void uErrorRef()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorRef",
                    (object)excel.Application.Run("uGetErrorRef")));
        }

        [Function()]
        public static object uGetErrorNum()
            => FunctionHost.Instance.ErrorNum;

        [Function()]
        public static bool uSetErrorNum([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorNum);

        [TestMethod]
        public void uErrorNum()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorNum",
                    (object)excel.Application.Run("uGetErrorNum")));
        }

        [Function()]
        public static object uGetErrorData()
            => FunctionHost.Instance.ErrorData;

        [Function()]
        public static bool uSetErrorData([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorData);

        [TestMethod]
        public void uErrorData()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorData",
                    (object)excel.Application.Run("uGetErrorData")));
        }

        [Function()]
        public static object uGetErrorValue()
            => FunctionHost.Instance.ErrorValue;

        [Function()]
        public static bool uSetErrorValue([Argument(Name = "[v1]")] object v1)
            => CompareError(v1, FunctionHost.Instance.ErrorValue);

        [TestMethod]
        public void uErrorValue()
        {
            using (var excel = new ExcelLoader())
                Assert.AreEqual(true, (bool)excel.Application.Run("uSetErrorValue",
                    (object)excel.Application.Run("uGetErrorValue")));
        }
    }
}
