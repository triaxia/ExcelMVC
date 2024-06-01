using ExcelMvc.Functions;
using Function.Interfaces;
using System;

namespace Samples
{
    public static class MarshalTests
    {
        [Function(Name = "uDouble")]
        public static double Double(double value)
        {
            throw new NotImplementedException();
            return value;
        }

        [Function(Name = "uFloat")]
        public static double Float(double value)
        {
            return value;
        }

        [Function(Name = "uDateTime")]
        public static double DatTime(double value)
        {
            return value;
        }

        [Function(Name = "uInt32")]
        public static int Int32(int value)
        {
            return value;
        }

        [Function(Name = "uUInt32")]
        public static uint UInt32(uint value)
        {
            return value;
        }

        [Function(Name = "uInt16")]
        public static short Int16(short value)
        {
            return value;
        }

        [Function(Name = "uUInt16")]
        public static ushort UInt16(ushort value)
        {
            return value;
        }

        [Function(Name = "uByte")]
        public static byte Byte(byte value)
        {
            return value;
        }

        [Function(Name = "uSByte")]
        public static sbyte SByte(sbyte value)
        {
            return value;
        }

        [Function(Name = "uBoolean")]
        public static bool Boolean(bool value)
        {
            return value;
        }

        [Function(Name = "uString")]
        public static string String(string value)
        {
            return value;
        }

        [Function(Name = "uDoubleArray")]
        public static double[] DoubleArray(double[] value)
        {
            return value;
        }

        [Function(Name = "uDoubleMatrix")]
        public static double[,] DoubleMatrix(double[,] value)
        {
            return value;
        }

        [Function(Name = "uInt32Array")]
        public static int[] Int32Array(int[] value)
        {
            return value;
        }

        [Function(Name = "uInt32Matrix")]
        public static int[,] Int32Matrix(int[,] value)
        {
            return value;
        }

        [Function(Name = "uDateTimeArray")]
        public static DateTime[] DateTimeArray(DateTime[] value)
        {
            return value;
        }

        [Function(Name = "uDateTimeMatrix")]
        public static DateTime[,] DateTimeMatrix(DateTime[,] value)
        {
            return value;
        }

        [Function(Name = "uStringArray")]
        public static string[] StringArray(string[] value)
        {
            return value;
        }

        [Function(Name = "uStringMatrix")]
        public static string[,] StringMatrix(string[,] value)
        {
            return value;
        }

        [Function(Name = "uObjectArray")]
        public static object[] ObjectArray(object[] value)
        {
            return value;
        }

        [Function(Name = "uObjectMatrix")]
        public static object[,] ObjectMatrix(object[,] value)
        {
            return value;
        }

        [Function(Name = "uObject")]
        public static object Object(object value)
        {
            return value;
        }

        [Function(Name = "uCaller")]
        public static string Caller()
        {
            return $"{Host.Instance.GetCallerReference()}";
        }

        [Function(Name = "uActiveSheetRangeValue")]
        public static object ActiveSheetRangeValue(int row, int column, int rowCount, int columnCount, object value)
        {
            var reference = Host.Instance.GetActivePageReference(row, column, rowCount, columnCount);
            Host.Instance.SetRangeValue(reference, value, true);
            return Host.Instance.GetRangeValue(reference);
        }

        [Function(Name = "uIsInFunctionWizard")]
        public static object IsInFunctionWizard(int a, int b, int c)
        {
            if (Host.Instance.IsInFunctionWizard())
                return "editing...";
            return a + b + c;
        }

        [Function(Name = "uHelp")]
        public static object uHelp(int a, int b, int c)
        {
            return "https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1";
        }

        [Function(Name = "uDefaultValue")]
        public static object uDefaultValue(int a, [Argument(Name = "[b]")] int b = 100, [Argument(Name = "[c]")] int c = 200)
        {
            return a + b + c;
        }

        [Function(Name = "uExceptionObject")]
        public static object uExceptionObject()
        {
            Host.Instance.ExceptionToFunctionResult = _=> Host.Instance.ErrorValue;
            throw new Exception(nameof(uExceptionObject));
        }

        [Function(Name = "uExceptionString")]
        public static string uExceptionString()
        {
            Host.Instance.ExceptionToFunctionResult = _ => Host.Instance.ErrorValue;
            throw new Exception(nameof(uExceptionString));
        }

        [Function(Name = "uExceptionInt")]
        public static int uExceptionInt()
        {
            Host.Instance.ExceptionToFunctionResult = _ => Host.Instance.ErrorValue;
            throw new Exception(nameof(uExceptionInt));
        }

        [Function(Name = "uExceptionMessage")]
        public static object uExceptionMessage()
        {
            Host.Instance.ExceptionToFunctionResult = e => $"{e}";
            throw new Exception(nameof(uExceptionInt));
        }
    }
}