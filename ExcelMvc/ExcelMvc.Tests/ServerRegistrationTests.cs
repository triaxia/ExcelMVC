using ExcelMvc.Rtd;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class ServerRegistrationTests
    {
        [TestMethod]
        public void Register()
        {
            var progId = ExcelMvc.Rtd.RtdRegistration.RegisterType(typeof(RtdServer));
            var type = Type.GetTypeFromProgID("ExcelMvc.Rtd", true);
            dynamic instance  = Activator.CreateInstance(type);
            Assert.AreEqual(5, instance.Add(2, 3));
        }
    }
}
