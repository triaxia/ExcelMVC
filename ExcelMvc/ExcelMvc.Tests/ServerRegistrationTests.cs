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
            RtdRegistration.PurgeProgIds();
            //RtdRegistration.RegisterType(typeof(Rtd101));
            //Activator.CreateInstance(type);
        }
    }
}
