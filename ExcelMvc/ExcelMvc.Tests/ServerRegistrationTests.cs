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
    [Guid("F80F202A-B862-4D50-AA51-F0481781CB4F")]
    [ComVisible(true)]
    [ProgId("ExcelMvc.TestServer")]
    public class TestServer
    {
        public int Add(int a, int b) => a + b;
    }


    [TestClass]
    public class ServerRegistrationTests
    {
        [TestMethod]
        public void Register()
        {
            var progId = ExcelMvc.Rtd.ServerRegistration.RegisterType(typeof(TestServer));
            var type = Type.GetTypeFromProgID(progId, true);
            dynamic instance  = Activator.CreateInstance(type);
            Assert.AreEqual(5, instance.Add(2, 3));
        }
    }
}
