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
    }
}