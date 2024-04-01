using ExcelMvc.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class XlMarshalTests
    {
        public static double AddDouble(double x, double y)
        {
            return x + y;
        }

        [TestMethod]
        public void MarshalSameType_AddTwoDouble()
        {
            var method = typeof(XlMarshalTests).GetMethod("AddDouble");
            var func = (FunctionDelegate.Function2) DelegateFactory.MakeOuterDelegate(method);

            var p1 = new XlMarshalContext();
            var p2 = new XlMarshalContext();
            var result = func(p1.DoubleToIntPtr(2), p2.DoubleToIntPtr(3));
            Assert.AreEqual(5.0, XlMarshalContext.IntPtrToDouble(result));
        }

        public static double AddMixed(double a, long b, int c, decimal d, byte e)
        {
            return a + (double)b + (double)c + (double)d + (double)e;
        }

        [TestMethod]
        public void MarshalMixedType_AddMany()
        {
            var method = typeof(XlMarshalTests).GetMethod("AddMixed");
            var func = (FunctionDelegate.Function5)DelegateFactory.MakeOuterDelegate(method);

            var p1 = new XlMarshalContext();
            var p2 = new XlMarshalContext();
            var p3 = new XlMarshalContext();
            var p4 = new XlMarshalContext();
            var p5 = new XlMarshalContext();
            var result = func(p1.DoubleToIntPtr(1), p2.LongToIntPtr(2), p3.IntToIntPtr(3), p4.DecimalToIntPtr(4), p5.ByteToIntPtr(5));
            Assert.AreEqual(15.0, XlMarshalContext.IntPtrToDouble(result));
        }

        public static string ConcatString(string a, string b)
        {
            return a + b;
        }

        [TestMethod]
        public void MarshalString_Concat()
        {
            var method = typeof(XlMarshalTests).GetMethod("ConcatString");
            var func = (FunctionDelegate.Function2)DelegateFactory.MakeOuterDelegate(method);

            var p1 = new XlMarshalContext();
            var p2 = new XlMarshalContext();
            var result = func(p1.StringToIntPtr("abc"), p2.StringToIntPtr("efg"));
            Assert.AreEqual("abcefg", XlMarshalContext.IntPtrToString(result));
        }
    }
}
