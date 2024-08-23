using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class FunctionTests
    {
        [TestMethod]
        public void uBool()
        {
            using (var excel = new ExcelLoader())
            {
                var result = (bool) excel.Application.Run("uBool", true);
                Assert.IsFalse(result);
            }
        }
    }
}