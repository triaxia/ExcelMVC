using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMvc
{
    public class TestRtd
    {
        public int Add(int a, int b) => a + b;
    }


    [TestClass]
    public class ServerRegistrationTests
    {
        [TestMethod]
        public void Register()
        {
            try
            {
                var progId = ExcelMvc.Rtd.ServerRegistration.Register(typeof(TestRtd));
                var type = Type.GetTypeFromProgID(progId, true);
                dynamic instance  = Activator.CreateInstance(type);
                Assert.AreEqual(5, instance.Add(2, 3));

            }
            catch (Exception ex)
            {
            }
        }
    }
}
