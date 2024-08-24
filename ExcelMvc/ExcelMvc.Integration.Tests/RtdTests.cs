using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class RtdTests
    {
        [Function(Name = "uTimer")]
        public static object uTimer(string name)
        {
            return FunctionHost.Instance.Rtd<TimerServer>(() => new TimerServer(), "", name);
        }

        [TestMethod]
        public void uTimer()
        {
            using (var excel = new ExcelLoader())
            {
                var items = new List<string>();
                var start = DateTime.UtcNow;
                while ((DateTime.UtcNow - start).TotalSeconds < 10)
                {
                    var result = (string)excel.Application.Run("uTimer", "test");
                    items.Add(result);
                }
                var distinct = items.Distinct().Count();
                Assert.IsTrue(distinct >= 10 && distinct <= 11);
            }
        }
    }
}