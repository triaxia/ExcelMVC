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
            return value;
        }

        [ExcelFunction(Name = "uDoubleMatrix")]
        public static double[,] DoubleMatrix(double[,] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uInt32Array")]
        public static int[] Int32Array(int[] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uInt32Matrix")]
        public static int[,] Int32Matrix(int[,] value)
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

        [ExcelFunction(Name = "uStringArray")]
        public static string[] StringArray(string[] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uStringMatrix")]
        public static string[,] StringMatrix(string[,] value)
        {
            return value;
        }

        [ExcelFunction(Name = "uObjectArray")]
        public static object[] ObjectArray(object[] value)
        {
            return value;
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

        [ExcelFunction(Name = "uCaller")]
        public static string Caller()
        {
            return $"{XlCall.GetCallerReference()}";
        }

        [ExcelFunction(Name = "uActiveSheetRangeValue")]
        public static object ActiveSheetRangeValue(int row, int column, int rowCount, int columnCount, object value)
        {
            var reference = XlCall.GetActiveSheetReference(row, column, rowCount, columnCount);
            reference.SetValue(value, true);
            return reference.GetValue();
        }

        [ExcelFunction(Name = "uIsInFunctionWizard")]
        public static object IsInFunctionWizard(int a, int b, int c)
        {
            if (XlCall.IsInFunctionWizard())
                return "editing...";
            return a + b + c;
        }
    }
}