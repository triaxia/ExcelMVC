using ExcelMvc.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using static ExcelMvc.Functions.FunctionDelegate;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class DelegateFactoryTests
    {
        public static double Add(double x, double y)
        {
            return x + y;
        }

        [TestMethod]
        public void Discover()
        {
            var method = typeof(DelegateFactoryTests).GetMethod("Add");
            var dele = DelegateFactory.MakeOuterDelegate(method);

            Function2 x = (Function2)dele;
            var d = x(IntPtr.Zero, IntPtr.Zero);
        }
    }
}
