using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class DefautTests
    {
        private const string DefaultString = "gu";
        private const int DefaultInt = int.MinValue / 2;
        private const double DefaultDouble = 123.456;

        [Function()]
        public static string uDefault(
            [Argument(Name = "[v1]")] string v1 = DefaultString,
            [Argument(Name = "[v2]")] int? v2 = DefaultInt,
            [Argument(Name = "[v3]")] double? v3 = DefaultDouble,
            [Argument(Name = "[v4]")] DateTime? v4 = null)
        {
            return $"{v1}|{v2}|{v3}|{v4}";
        }

        [TestMethod]
        public void uDefault()
        {
            using (var excel = new ExcelLoader())
            {
                //var result = (string)excel.Application.Run("uDefault");
                //Assert.AreEqual($"{DefaultString}|{DefaultInt}|{DefaultDouble}|{DefaultDate}", result);

                var today = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var result = (string)excel.Application.Run("uDefault", "a", 1, 2, today);
                Assert.AreEqual($"a|1|2|{today}", result);
            }
        }
    }
}