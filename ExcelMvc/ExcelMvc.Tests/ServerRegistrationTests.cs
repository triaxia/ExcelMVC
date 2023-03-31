using ExcelMvc.Rtd;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class ServerRegistrationTests
    {
        [TestMethod]
        public void Register()
        {
            var progId = ExcelMvc.Rtd.ServerRegistration.RegisterType(typeof(TestServer));
            var type = Type.GetTypeFromProgID("ExcelMvc.TestServer", true);
            dynamic instance  = Activator.CreateInstance(type);
            Assert.AreEqual(5, instance.Add(2, 3));
        }
    }
}
