using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class ObjectArrayTests
    {
        [Function()]
        public static object[] uObjectArray(object[] v1, object[] v2 = null)
        {
            return v2 == null ? v1 : v1.Concat(v2).ToArray();
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

                jagged = (Array)(object)excel.Application.Run("uObjectArray", cells, cells);
                result = new object[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(today, DateTime.Parse((string)result[0]));
                Assert.AreEqual(int.MaxValue, (double)result[1]);
                Assert.AreEqual(double.MaxValue, (double)result[2]);
                Assert.AreEqual(today, DateTime.Parse((string)result[3]));
                Assert.AreEqual(int.MaxValue, (double)result[4]);
                Assert.AreEqual(double.MaxValue, (double)result[5]);
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
                Assert.AreEqual(int.MaxValue, (double)result[0, 1]);
                Assert.AreEqual(double.MaxValue, (double)result[0, 2]);
                Assert.AreEqual(string.Empty, (string)result[1, 0]);
                Assert.AreEqual(short.MaxValue, (double)result[1, 1]);
                Assert.AreEqual(float.MaxValue, (double)result[1, 2]);
            }
        }

        [Function()]
        public static object[] uConcatObjectArray(object[] v1, [Argument(Name = "[v2]")] object[] v2 = null)
        {
            return v2 == null ? v1 : v1.Concat(v2).ToArray();
        }

        [TestMethod]
        public void uConcatObjectArray()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new double[] { 1, 2, 3 };

                var jagged = (Array)(object)excel.Application.Run("uConcatIntArray", cells);
                var result = new double[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(1, result[0]);
                Assert.AreEqual(2, result[1]);
                Assert.AreEqual(3, result[2]);

                jagged = (Array)(object)excel.Application.Run("uConcatIntArray", cells, cells);
                result = new double[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(1, result[0]);
                Assert.AreEqual(2, result[1]);
                Assert.AreEqual(3, result[2]);
                Assert.AreEqual(1, result[3]);
                Assert.AreEqual(2, result[4]);
                Assert.AreEqual(3, result[5]);
            }
        }

        [Function()]
        public static object[,] uConcatObjectMatrix(object[,] v1, [Argument(Name = "[v2]")] object[,] v2 = null)
        {
            if (v2 == null) return v1;

            var result = Array.CreateInstance(typeof(object), v1.GetLength(0) + v2.GetLength(0), v1.GetLength(1));
            for (int i = 0; i < v1.GetLength(0); i++)
                for (int j = 0; j < v1.GetLength(1); j++)
                    result.SetValue(v1[i, j], i, j);
            for (int i = 0; i < v2.GetLength(0); i++)
                for (int j = 0; j < v2.GetLength(1); j++)
                    result.SetValue(v2[i, j], i + v1.GetLength(0), j);
            return (object[,])result;
        }

        [TestMethod]
        public void uConcatObjectMatrix()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new[,] { { 1, 2, 3 }, { 4, 5, 6 } };

                var jagged = (Array)(object)excel.Application.Run("uConcatObjectMatrix", cells);
                var result = new double[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(1, result[0, 0]);
                Assert.AreEqual(2, result[0, 1]);
                Assert.AreEqual(3, result[0, 2]);
                Assert.AreEqual(4, result[1, 0]);
                Assert.AreEqual(5, result[1, 1]);
                Assert.AreEqual(6, result[1, 2]);

                jagged = (Array)(object)excel.Application.Run("uConcatObjectMatrix", cells, cells);
                result = new double[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(1, result[0, 0]);
                Assert.AreEqual(2, result[0, 1]);
                Assert.AreEqual(3, result[0, 2]);
                Assert.AreEqual(4, result[1, 0]);
                Assert.AreEqual(5, result[1, 1]);
                Assert.AreEqual(6, result[1, 2]);
                Assert.AreEqual(1, result[2, 0]);
                Assert.AreEqual(2, result[2, 1]);
                Assert.AreEqual(3, result[2, 2]);
                Assert.AreEqual(4, result[3, 0]);
                Assert.AreEqual(5, result[3, 1]);
                Assert.AreEqual(6, result[3, 2]);
            }
        }

        [Function()]
        public static object[] uObjectArraySingleValue([Argument(Name = "[v1]")] object[] v1 = null)
        {
            return v1;
        }


        [TestMethod]
        public void uObjectArraySingleValue()
        {
            using (var excel = new ExcelLoader())
            {
                var value = Guid.NewGuid().ToString();
                var jagged = (Array)(object)excel.Application.Run("uObjectArraySingleValue", value);
                var result = new object[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(value, (string)result[0]);
            }
        }
    }
}