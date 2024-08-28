using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class StringArrayTests
    {
        [Function()]
        public static string[] uStringArray(string[] v1, [Argument(Name = "[v2]")] string v2 = "")
        {
            for (int i = 0; i < v1.Length; i++)
                v1[i] = $"{v1[i]}{v2}";
            return v1;
        }

        [TestMethod]
        public void uStringArray()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new string[] { "1", "2", "3" };

                var jagged = (Array)(object)excel.Application.Run("uStringArray", cells);
                var result = new string[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("1", result[0]);
                Assert.AreEqual("2", result[1]);
                Assert.AreEqual("3", result[2]);

                jagged = (Array)(object)excel.Application.Run("uStringArray", cells, "10");
                result = new string[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("110", result[0]);
                Assert.AreEqual("210", result[1]);
                Assert.AreEqual("310", result[2]);
            }
        }

        [Function()]
        public static string[,] uStringMatrix(string[,] v1, [Argument(Name = "[v2]")] string v2 = "")
        {
            for (int i = 0; i < v1.GetLength(0); i++)
                for (int j = 0; j < v1.GetLength(1); j++)
                    v1[i, j] = $"{v1[i, j]}{v2}";
            return v1;
        }

        [TestMethod]
        public void uStringMatrix()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new string[,] { { "1", "2", "3" }, { "4", "5", "6" } };

                var jagged = (Array)(object)excel.Application.Run("uStringMatrix", cells);
                var result = new string[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("1", result[0, 0]);
                Assert.AreEqual("2", result[0, 1]);
                Assert.AreEqual("3", result[0, 2]);
                Assert.AreEqual("4", result[1, 0]);
                Assert.AreEqual("5", result[1, 1]);
                Assert.AreEqual("6", result[1, 2]);

                jagged = (Array)(object)excel.Application.Run("uStringMatrix", cells, "10");
                result = new string[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("110", result[0, 0]);
                Assert.AreEqual("210", result[0, 1]);
                Assert.AreEqual("310", result[0, 2]);
                Assert.AreEqual("410", result[1, 0]);
                Assert.AreEqual("510", result[1, 1]);
                Assert.AreEqual("610", result[1, 2]);
            }
        }

        [Function()]
        public static string[] uConcatStringArray(string[] v1, [Argument(Name = "[v2]")] string[] v2 = null)
        {
            return v2 == null ? v1 : v1.Concat(v2).ToArray();
        }

        [TestMethod]
        public void uConcatStringArray()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new string[] { "1", "2", "3" };

                var jagged = (Array)(object)excel.Application.Run("uConcatStringArray", cells);
                var result = new string[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("1", result[0]);
                Assert.AreEqual("2", result[1]);
                Assert.AreEqual("3", result[2]);

                jagged = (Array)(object)excel.Application.Run("uConcatStringArray", cells, cells);
                result = new string[jagged.Length];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("1", result[0]);
                Assert.AreEqual("2", result[1]);
                Assert.AreEqual("3", result[2]);
                Assert.AreEqual("1", result[3]);
                Assert.AreEqual("2", result[4]);
                Assert.AreEqual("3", result[5]);
            }
        }

        [Function()]
        public static string[,] uConcatStringMatrix(string[,] v1, [Argument(Name = "[v2]")] string[,] v2 = null)
        {
            if (v2 == null) return v1;

            var result = Array.CreateInstance(typeof(string), v1.GetLength(0) + v2.GetLength(0), v1.GetLength(1));
            for (int i = 0; i < v1.GetLength(0); i++)
                for (int j = 0; j < v1.GetLength(1); j++)
                    result.SetValue(v1[i, j], i, j);
            for (int i = 0; i < v2.GetLength(0); i++)
                for (int j = 0; j < v2.GetLength(1); j++)
                    result.SetValue(v2[i, j], i + v1.GetLength(0), j);
            return (string[,])result;
        }

        [TestMethod]
        public void uConcatStringMatrix()
        {
            using (var excel = new ExcelLoader())
            {
                var cells = new string[,] { { "1", "2", "3" }, { "4", "5", "6" } };

                var jagged = (Array)(object)excel.Application.Run("uConcatStringMatrix", cells);
                var result = new string[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("1", result[0, 0]);
                Assert.AreEqual("2", result[0, 1]);
                Assert.AreEqual("3", result[0, 2]);
                Assert.AreEqual("4", result[1, 0]);
                Assert.AreEqual("5", result[1, 1]);
                Assert.AreEqual("6", result[1, 2]);

                jagged = (Array)(object)excel.Application.Run("uConcatStringMatrix", cells, cells);
                result = new string[jagged.GetLength(0), jagged.GetLength(1)];
                Array.Copy(jagged, result, result.Length);
                Assert.AreEqual("1", result[0, 0]);
                Assert.AreEqual("2", result[0, 1]);
                Assert.AreEqual("3", result[0, 2]);
                Assert.AreEqual("4", result[1, 0]);
                Assert.AreEqual("5", result[1, 1]);
                Assert.AreEqual("6", result[1, 2]);
                Assert.AreEqual("1", result[2, 0]);
                Assert.AreEqual("2", result[2, 1]);
                Assert.AreEqual("3", result[2, 2]);
                Assert.AreEqual("4", result[3, 0]);
                Assert.AreEqual("5", result[3, 1]);
                Assert.AreEqual("6", result[3, 2]);
            }
        }

    }
}