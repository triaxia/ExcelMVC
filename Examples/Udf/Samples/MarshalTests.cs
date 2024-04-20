using ExcelMvc.Functions;
using System;

namespace Samples
{
    public static class MarshalTests
    {
        [ExcelFunction(Name = "uDouble")]
        public static double Double(double value)
        {
            return value;
        }

        [ExcelFunction(Name = "uFloat")]
        public static double Float(double value)
        {
            return value;
        }

        [ExcelFunction(Name = "uDateTime")]
        public static double DatTime(double value)
        {
            return value;
        }

        [ExcelFunction(Name = "uInt32")]
        public static int Int32(int value)
        {
            return value;
        }

        [ExcelFunction(Name = "uUInt32")]
        public static uint UInt32(uint value)
        {
            return value;
        }

        [ExcelFunction(Name = "uInt16")]
        public static short Int16(short value)
        {
            return value;
        }

        [ExcelFunction(Name = "uUInt16")]
        public static ushort UInt16(ushort value)
        {
            return value;
        }

        [ExcelFunction(Name = "uByte")]
        public static byte Byte(byte value)
        {
            return value;
        }

        [ExcelFunction(Name = "uSByte")]
        public static sbyte SByte(sbyte value)
        {
            return value;
        }

        [ExcelFunction(Name = "uBoolean")]
        public static bool Boolean(bool value)
        {
            return value;
        }

        [ExcelFunction(Name = "uString")]
        public static string String(string value)
        {
            return value;
        }

        [ExcelFunction(Name = "uDoubleArray")]
        public static double[] DoubleArray(double[] value)
        {
            return new double [] { };
        }

        [ExcelFunction(Name = "uDoubleMatrix")]
        public static double[,] DoubleMatrix(double[,] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uDateTimeArray")]
        public static DateTime[] DateTimeArray(DateTime[] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uDateTimeMatrix")]
        public static DateTime[,] DateTimeMatrix(DateTime[,] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uObjectArray")]
        public static object[] ObjectArray(object[] value)
        {
            return new object[] { };
        }

        [ExcelFunction(Name = "uObjectMatrix")]
        public static object[,] ObjectMatrix(object[,] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uObject")]
        public static object Object(object value)
        {
            return value;
        }
    }
}