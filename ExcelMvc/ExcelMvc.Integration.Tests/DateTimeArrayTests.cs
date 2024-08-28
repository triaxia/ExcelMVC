using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Linq;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class DateTimeArrayTests
    {
        [Function()]
        public static DateTime[] uDateTimeArray(DateTime[] v1, [Argument(Name = "[v2]")] int? v2 = 0)
        {
            for (int i = 0; i < v1.Length; i++)
                v1[i] = v1[i].AddDays(v2.Value);
            return v1;
        }

        [TestMethod]
        public void uDateTimeArray()
        {
            using (var excel = new ExcelLoader())
            {
                var d0 = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var d1 = d0.AddDays(1); 
                var d2 = d0.AddDays(2);
                var cells = new DateTime[] { d0, d1, d2 };

                var jagged = (Array)(object)excel.Application.Run("uDateTimeArray", cells);
                var result = new double[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(d0, DateTime.FromOADate(result[0]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[1]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[2]));
            }
        }

        [Function()]
        public static DateTime[,] uDateTimeMatrix(DateTime[,] v1, [Argument(Name = "[v2]")] int? v2 = 0)
        {
            for (int i = 0; i < v1.GetLength(0); i++)
                for (int j = 0; j < v1.GetLength(1); j++)
                    v1[i, j] = v1[i, j].AddDays(v2.Value);
            return v1;
        }

        [TestMethod]
        public void uDateTimeMatrix()
        {
            using (var excel = new ExcelLoader())
            {
                var d0 = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var d1 = d0.AddDays(1);
                var d2 = d0.AddDays(2);
                var d3 = d0.AddDays(3);
                var d4 = d0.AddDays(4);
                var d5 = d0.AddDays(5);
                var cells = new DateTime[,] { { d0, d1, d2 }, { d3, d4, d5 } };

                var jagged = (Array)(object)excel.Application.Run("uDateTimeMatrix", cells);
                var result = new double[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(d0, DateTime.FromOADate(result[0, 0]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[0, 1]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[0, 2]));
                Assert.AreEqual(d3, DateTime.FromOADate(result[1, 0]));
                Assert.AreEqual(d4, DateTime.FromOADate(result[1, 1]));
                Assert.AreEqual(d5, DateTime.FromOADate(result[1, 2]));
            }
        }

        [Function()]
        public static DateTime[] uConcatDateTimeArray(DateTime[] v1, [Argument(Name = "[v2]")] DateTime[] v2 = null)
        {
            return v2 == null ? v1 : v1.Concat(v2).ToArray();
        }

        [TestMethod]
        public void uConcatDateTimeArray()
        {
            using (var excel = new ExcelLoader())
            {
                var d0 = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var d1 = d0.AddDays(1);
                var d2 = d0.AddDays(2);
                var cells = new DateTime[] { d0, d1, d2 };

                var jagged = (Array)(object)excel.Application.Run("uConcatDateTimeArray", cells);
                var result = new double[jagged.GetLength(0)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(d0, DateTime.FromOADate(result[0]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[1]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[2]));

                jagged = (Array)(object)excel.Application.Run("uConcatDateTimeArray", cells, cells);
                result = new double[jagged.GetLength(0)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(d0, DateTime.FromOADate(result[0]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[1]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[2]));
                Assert.AreEqual(d0, DateTime.FromOADate(result[3]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[4]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[5]));
            }
        }

        [Function()]
        public static DateTime[,] uConcatDateTimeMatrix(DateTime[,] v1, [Argument(Name = "[v2]")] DateTime[,] v2 = null)
        {
            if (v2 == null) return v1;

            var result = Array.CreateInstance(typeof(DateTime), v1.GetLength(0) + v2.GetLength(0), v1.GetLength(1));
            for (int i = 0; i < v1.GetLength(0); i++)
                for (int j = 0; j < v1.GetLength(1); j++)
                    result.SetValue(v1[i, j], i, j);
            for (int i = 0; i < v2.GetLength(0); i++)
                for (int j = 0; j < v2.GetLength(1); j++)
                    result.SetValue(v2[i, j], i + v1.GetLength(0), j);
            return (DateTime[,])result;
        }

        [TestMethod]
        public void uConcatDateTimeMatrix()
        {
            using (var excel = new ExcelLoader())
            {
                var d0 = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var d1 = d0.AddDays(1);
                var d2 = d0.AddDays(2);
                var d3 = d0.AddDays(3);
                var d4 = d0.AddDays(4);
                var d5 = d0.AddDays(5);
                var cells = new DateTime[,] { { d0, d1, d2 }, { d3, d4, d5 } };

                var jagged = (Array)(object)excel.Application.Run("uConcatDateTimeMatrix", cells);
                var result = new double[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(d0, DateTime.FromOADate(result[0, 0]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[0, 1]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[0, 2]));
                Assert.AreEqual(d3, DateTime.FromOADate(result[1, 0]));
                Assert.AreEqual(d4, DateTime.FromOADate(result[1, 1]));
                Assert.AreEqual(d5, DateTime.FromOADate(result[1, 2]));

                jagged = (Array)(object)excel.Application.Run("uConcatDateTimeMatrix", cells, cells);
                result = new double[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual(d0, DateTime.FromOADate(result[0, 0]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[0, 1]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[0, 2]));
                Assert.AreEqual(d3, DateTime.FromOADate(result[1, 0]));
                Assert.AreEqual(d4, DateTime.FromOADate(result[1, 1]));
                Assert.AreEqual(d5, DateTime.FromOADate(result[1, 2]));
                Assert.AreEqual(d0, DateTime.FromOADate(result[2, 0]));
                Assert.AreEqual(d1, DateTime.FromOADate(result[2, 1]));
                Assert.AreEqual(d2, DateTime.FromOADate(result[2, 2]));
                Assert.AreEqual(d3, DateTime.FromOADate(result[3, 0]));
                Assert.AreEqual(d4, DateTime.FromOADate(result[3, 1]));
                Assert.AreEqual(d5, DateTime.FromOADate(result[3, 2]));
            }
        }
    }
}