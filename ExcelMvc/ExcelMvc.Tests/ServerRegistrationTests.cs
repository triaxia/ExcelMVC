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
            //var progId = ExcelMvc.Rtd.RtdRegistration.RegisterType(typeof(RtdServerFactory.RtdServer001));
            //var type = Type.GetTypeFromProgID("ExcelMvc.Rtd00", true);
            //dynamic instance  = Activator.CreateInstance(type);
            //Assert.IsNotNull(instance);
            //ExcelMvc.Rtd.RtdRegistration.DeleteProgId("ExcelMvc.Rtd00");
            var dd = RtdServerFactory.Create(new RtdServerImplTest());
        }
    }
}
