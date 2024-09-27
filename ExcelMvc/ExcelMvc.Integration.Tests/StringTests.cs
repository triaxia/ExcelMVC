using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class StringTests
    {
        [Function()]
        public static string uStringOptionalWithNoDefault([Argument(Name = "[v1]")] string v1)
        {
            return v1;
        }

        [TestMethod]
        public void uStringOptionalWithNoDefault()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (object)excel.Application.Run("uStringOptionalWithNoDefault");
                Assert.AreEqual("", result);
                
                var value = Guid.NewGuid().ToString();  
                result = (string) (object)excel.Application.Run("uStringOptionalWithNoDefault", value);
                Assert.AreEqual(value, result);
            }
        }
    }
}
