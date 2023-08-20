using ExcelMvc.Rtd;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class ServerRegistrationTests
    {
        [TestMethod]
        public void Register()
        {
            var type = Type.GetTypeFromProgID("ExcelMvc.Rtd001");
            Activator.CreateInstance(type);
        }
    }
}
