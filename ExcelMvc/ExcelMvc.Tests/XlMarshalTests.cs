using ExcelMvc.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class XlMarshalTests
    {
        public static double MarshalDouble(double x) => x;
        [TestMethod]
        public void Marshal_Double()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDouble));
            AssertMarshal<Double>(method, double.MaxValue);
            AssertMarshal<Double>(method, double.MinValue);
        }

        public static bool MarshalBoolean(bool x) => x;
        [TestMethod]
        public void Marshal_Boolean()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalBoolean));
            AssertMarshal<bool>(method, true);
        }

        public static DateTime MarshalDateTime(DateTime x) => x;
        [TestMethod]
        public void Marshal_DateTime()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDateTime));
            var max = DateTime.FromOADate(DateTime.MaxValue.ToOADate());
            AssertMarshal<DateTime>(method, max);
            var min = DateTime.FromOADate(DateTime.MinValue.ToOADate());
            AssertMarshal<DateTime>(method, min);
        }

        public static float MarshalSingle(float x) => x;
        [TestMethod]
        public void Marshal_Single()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalSingle));
            AssertMarshal<float>(method, float.MaxValue);
            AssertMarshal<float>(method, float.MinValue);
        }

        public static int MarshalInt32(int x) => x;
        [TestMethod]
        public void Marshal_Int32()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalInt32));
            AssertMarshal<int>(method, int.MaxValue);
            AssertMarshal<int>(method, int.MinValue);
        }

        public static uint MarshalUInt32(uint x) => x;
        [TestMethod]
        public void Marshal_UInt32()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalUInt32));
            AssertMarshal<uint>(method, uint.MaxValue);
            AssertMarshal<uint>(method, uint.MinValue);
        }

        public static short MarshalInt16(short x) => x;
        [TestMethod]
        public void Marshal_Int16()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalInt16));
            AssertMarshal<short>(method, short.MaxValue);
            AssertMarshal<short>(method, short.MinValue);
        }

        public static ushort MarshalUInt16(ushort x) => x;
        [TestMethod]
        public void Marshal_UInt16()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalUInt16));
            AssertMarshal<ushort>(method, ushort.MaxValue);
            AssertMarshal<ushort>(method, ushort.MinValue);
        }

        public static byte MarshalByte(byte x) => x;
        [TestMethod]
        public void Marshal_Byte()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalByte));
            AssertMarshal<byte>(method, byte.MaxValue);
            AssertMarshal<byte>(method, byte.MinValue);
        }

        public static sbyte MarshalSByte(sbyte x) => x;
        [TestMethod]
        public void Marshal_SByte()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalSByte));
            AssertMarshal<sbyte>(method, sbyte.MaxValue);
            AssertMarshal<sbyte>(method, sbyte.MinValue);
        }

        public static string MarshalString(string x) => x;
        [TestMethod]
        public void Marshal_String()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalString));
            AssertMarshal<string>(method, Guid.NewGuid().ToString());
        }

        private static void AssertMarshal<TValue>(MethodInfo method, TValue value)
        {
            var func = (FunctionDelegate.Function1)DelegateFactory.MakeOuterDelegate(method);
            var name = typeof(TValue).Name;
            var p1 = new XlMarshalContext();
            var incoming = p1.GetType().GetMethod($"{name}ToIntPtr");

            var result = func((IntPtr)incoming.Invoke(p1, new object[] { value }));
            var outgoing = typeof(XlMarshalContext).GetMethod($"IntPtrTo{name}");
            Assert.AreEqual(value, outgoing.Invoke(null, new object[] { result }));
        }
    }
}
