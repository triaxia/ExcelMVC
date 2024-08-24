using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class DoubleArrayTests
    {
        [Function()]
        public static double[] uDoubleArray(double[] v1, [Argument(Name = "[v2]")] double? v2 = 1)
        {
            for (int i = 0; i < v1.Length; i++)
                v1[i] = v1[i] * v2.Value;
            return v1;
        }

        [TestMethod]
        public void uDoubleArray()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new double[] { 1, 2, 3 };

                var jagged = (Array)(object)excel.Application.Run("uDoubleArray", cells);
                var result = new double[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(1, result[0]);
                Assert.AreEqual(2, result[1]);
                Assert.AreEqual(3, result[2]);

                jagged = (Array)(object)excel.Application.Run("uDoubleArray", cells, 10);
                result = new double[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(10, result[0]);
                Assert.AreEqual(20, result[1]);
                Assert.AreEqual(30, result[2]);
            }
        }

        [Function()]
        public static double[,] uDoubleMatrix(double[,] v1, [Argument(Name = "[v2]")] double? v2 = 1)
        {
            for (int i = 0; i < v1.GetLength(0); i++)
                for (int j = 0; j < v1.GetLength(1); j++)
                    v1[i, j] = v1[i, j] * v2.Value;
            return v1;
        }

        [TestMethod]
        public void uDoubleMatrix()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new double[,] { { 1, 2, 3 }, { 4, 5, 6 } };

                var jagged = (Array)(object)excel.Application.Run("uDoubleMatrix", cells);
                var result = new double[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(1, result[0, 0]);
                Assert.AreEqual(2, result[0, 1]);
                Assert.AreEqual(3, result[0, 2]);
                Assert.AreEqual(4, result[1, 0]);
                Assert.AreEqual(5, result[1, 1]);
                Assert.AreEqual(6, result[1, 2]);

                jagged = (Array)(object)excel.Application.Run("uDoubleMatrix", cells, 10);
                result = new double[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(10, result[0, 0]);
                Assert.AreEqual(20, result[0, 1]);
                Assert.AreEqual(30, result[0, 2]);
                Assert.AreEqual(40, result[1, 0]);
                Assert.AreEqual(50, result[1, 1]);
                Assert.AreEqual(60, result[1, 2]);
            }
        }
    }
}