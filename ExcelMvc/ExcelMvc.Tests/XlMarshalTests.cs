using ExcelMvc.Functions;
using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Reflection;

namespace ExcelMvc.Tests
{
    [TestClass]
    public class XlMarshalTests
    {
        public XlMarshalTests()
        {
            Host.Instance = new ExcelFunctionHost();
        }

        public static double MarshalDouble(double x) => x;
        [TestMethod]
        public void Marshal_Double()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDouble));
            AssertMarshal(method, double.MaxValue);
            AssertMarshal(method, double.MinValue);
        }

        public static bool MarshalBoolean(bool x) => x;
        [TestMethod]
        public void Marshal_Boolean()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalBoolean));
            AssertMarshal(method, true);
        }

        public static DateTime MarshalDateTime(DateTime x) => x;
        [TestMethod]
        public void Marshal_DateTime()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDateTime));
            var max = DateTime.FromOADate(DateTime.MaxValue.ToOADate());
            AssertMarshal(method, max);
            var min = DateTime.FromOADate(DateTime.MinValue.ToOADate());
            AssertMarshal(method, min);
        }

        public static float MarshalSingle(float x) => x;
        [TestMethod]
        public void Marshal_Single()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalSingle));
            AssertMarshal(method, float.MaxValue);
            AssertMarshal(method, float.MinValue);
        }

        public static int MarshalInt32(int x) => x;
        [TestMethod]
        public void Marshal_Int32()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalInt32));
            AssertMarshal(method, int.MaxValue);
            AssertMarshal(method, int.MinValue);
        }

        public static uint MarshalUInt32(uint x) => x;
        [TestMethod]
        public void Marshal_UInt32()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalUInt32));
            AssertMarshal(method, uint.MaxValue);
            AssertMarshal(method, uint.MinValue);
        }

        public static short MarshalInt16(short x) => x;
        [TestMethod]
        public void Marshal_Int16()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalInt16));
            AssertMarshal(method, short.MaxValue);
            AssertMarshal(method, short.MinValue);
        }

        public static ushort MarshalUInt16(ushort x) => x;
        [TestMethod]
        public void Marshal_UInt16()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalUInt16));
            AssertMarshal(method, ushort.MaxValue);
            AssertMarshal(method, ushort.MinValue);
        }

        public static byte MarshalByte(byte x) => x;
        [TestMethod]
        public void Marshal_Byte()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalByte));
            AssertMarshal(method, byte.MaxValue);
            AssertMarshal(method, byte.MinValue);
        }

        public static sbyte MarshalSByte(sbyte x) => x;
        [TestMethod]
        public void Marshal_SByte()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalSByte));
            AssertMarshal(method, sbyte.MaxValue);
            AssertMarshal(method, sbyte.MinValue);
        }

        public static string MarshalString(string x) => x;
        [TestMethod]
        public void Marshal_String()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalString));
            AssertMarshal(method, Guid.NewGuid().ToString());
        }

        public static double[] MarshalDoubleArray(double[] x) => x;
        [TestMethod]
        public void Marshal_DoubleArray()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDoubleArray));
            AssertMarshal(method, new double[] { 1, 2, 3, 4 }, true);
        }

        public static DateTime[] MarshalDateTimeArray(DateTime[] x) => x;
        [TestMethod]
        public void Marshal_DateTimeArray()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDateTimeArray));
            var start = DateTime.FromOADate(DateTime.Now.ToOADate());
            AssertMarshal(method, new DateTime[] { start.AddDays(1), start.AddDays(2), start.AddDays(3)}, true);
        }

        public static double[,] MarshalDoubleMatrix(double[,] x) => x;
        [TestMethod]
        public void Marshal_DoubleMatrix()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDoubleMatrix));
            AssertMarshal(method, new double[,] { { 11, 12, 13, 14 }, { 21, 22, 23, 24 } }, true);
        }

        public static DateTime[,] MarshalDateTimeMatrix(DateTime[,] x) => x;
        [TestMethod]
        public void Marshal_DateTimeMatrix()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalDateTimeMatrix));
            var start = DateTime.FromOADate(DateTime.Now.ToOADate());
            AssertMarshal(method, new DateTime[,] { 
                { start.AddDays(11), start.AddDays(12), start.AddDays(13) },
                { start.AddDays(21), start.AddDays(22), start.AddDays(23) } }, true);
        }

        public static object[] MarshalObjectArray(object[] x) => x;
        [TestMethod]
        public void Marshal_ObjectArray()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalObjectArray));
            AssertMarshal(method, new object[] { "a", "b", "c", 1, 2, 3 }, true);
        }

        public static object[,] MarshalObjectMatrix(object[,] x) => x;
        [TestMethod]
        public void Marshal_ObjectMatrix()
        {
            var method = typeof(XlMarshalTests).GetMethod(nameof(MarshalObjectMatrix));
            AssertMarshal(method, new object[,] { { "a", "b", "c" }, { 1, 2, 3 } }, true);
        }

        private static void AssertMarshal<TValue>(MethodInfo method, TValue value
            ,bool isCollection = false)
        {
            var func = (FunctionDelegate.Function1)DelegateFactory.MakeOuterDelegate(method
                , new FunctionDefinition { Name="Test"});
            var name = typeof(TValue).Name.Replace("[]", "Array").Replace("[,]", "Matrix");
            var p1 = new XlMarshalContext();
            var incoming = p1.GetType().GetMethod($"{name}ToIntPtr");

            var inner = func((IntPtr)incoming.Invoke(p1, new object[] { value }));
            var outgoing = typeof(XlMarshalContext).GetMethod($"IntPtrTo{name}");
            var outer = outgoing.Invoke(null, new object[] { inner, null, false });
            if (isCollection)
                CollectionAssert.AreEqual((ICollection)value, (ICollection)outer);
            else
                Assert.AreEqual(value, outer);
        }
    }
}
