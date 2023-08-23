using ExcelMvc.Rtd;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class ServerRegistrationTests
    {
        [TestMethod]
        public void Register()
        {
            RtdRegistration.RegisterType(typeof(Rtd001));
            //Activator.CreateInstance(type);
        }
    }
}
