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
            RtdRegistration.RegisterType(typeof(Rtd002));
            //Activator.CreateInstance(type);
        }
    }
}
