using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class DateTimeTests
    {
        [Function()]
        public static DateTime uDateTime(DateTime v1, [Argument(Name = "[v2]")] int? v2 = 0)
        {
            return v1.AddDays(v2.Value);
        }

        [TestMethod]
        public void uDateTime()
        {
            using (var excel = new ExcelLoader())
            {
                var today = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var result = DateTime.FromOADate(excel.Application.Run("uDateTime", today));
                Assert.AreEqual(today, result);
                result = DateTime.FromOADate(excel.Application.Run("uDateTime", today, 356));
                Assert.AreEqual(today.AddDays(356), result);
            }
        }

        [Function()]
        public static DateTime uDateTimeDiff(DateTime v1, [Argument(Name = "[v2]")] DateTime? v2 = null)
        {
            return v1.AddDays(v2 == null ? 0  : (v2.Value - v1).TotalDays);
        }

        [TestMethod]
        public void uDateTimeDiff()
        {
            using (var excel = new ExcelLoader())
            {
                var today = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var result = DateTime.FromOADate(excel.Application.Run("uDateTimeDiff", today));
                Assert.AreEqual(today, result);
                result = DateTime.FromOADate(excel.Application.Run("uDateTimeDiff", today, today.AddDays(2)));
                Assert.AreEqual(today.AddDays(2), result);
            }
        }

        [Function()]
        public static double uDateTimeDiffDefault(DateTime v1 = default, DateTime v2 = default)
        {
            return  (v1 - v2).TotalDays;
        }

        [TestMethod]
        public void uDateTimeDiffDefault()
        {
            using (var excel = new ExcelLoader())
            {
                var today = DateTime.SpecifyKind(DateTime.Today, DateTimeKind.Unspecified);
                var result = excel.Application.Run("uDateTimeDiffDefault", today);
                Assert.AreEqual(today.ToOADate(), result);

                result = excel.Application.Run("uDateTimeDiffDefault", today, today.AddDays(13));
                Assert.AreEqual(-13, result);
            }
        }

    }
}