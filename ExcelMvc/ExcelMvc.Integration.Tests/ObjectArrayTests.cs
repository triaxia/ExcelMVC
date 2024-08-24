using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class ObjectArrayTests
    {
        [Function()]
        public static object[] uObjectArray(object[] v1)
        {
            return v1;
        }

        [TestMethod]
        public void uObjectArray()
        {
            using (var excel = new ExcelLoader())
            {
                var today = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var cells = new object[] { today, int.MaxValue, double.MaxValue };

                var jagged = (Array)(object)excel.Application.Run("uObjectArray", cells);
                var result = new object[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(today, DateTime.Parse((string)result[0]));
                Assert.AreEqual(int.MaxValue, (double)result[1]);
                Assert.AreEqual(double.MaxValue, (double)result[2]);
            }
        }

        [Function()]
        public static object[,] uObjectMatrix(object[,] v1)
        {
            return v1;
        }

        [TestMethod]
        public void uObjectMatrix()
        {
            using (var excel = new ExcelLoader())
            {
                var today = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var cells = new object[,] { { today, int.MaxValue, double.MaxValue }, { string.Empty, short.MaxValue, float.MaxValue } };

                var jagged = (Array)(object)excel.Application.Run("uObjectMatrix", cells);
                var result = new object[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(today, DateTime.Parse((string)result[0, 0]));
                Assert.AreEqual(int.MaxValue,(double)result[0, 1]);
                Assert.AreEqual(double.MaxValue, (double)result[0, 2]);
                Assert.AreEqual(string.Empty, (string)result[1, 0]);
                Assert.AreEqual(short.MaxValue, (double)result[1, 1]);
                Assert.AreEqual(float.MaxValue, (double)result[1, 2]);
            }
        }
    }
}