using Mvc;
using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class XlCall
    {
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "RegisterFunction")]
        public static extern IntPtr RegisterFunction64(IntPtr function);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "RegisterFunction")]
        public static extern IntPtr RegisterFunction32(IntPtr function);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "AsyncReturn")]
        public static extern IntPtr AsyncReturn64(IntPtr handle, IntPtr result);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "AsyncReturn")]
        public static extern IntPtr AsyncReturn32(IntPtr handle, IntPtr result);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "xlAutoFree12")]
        public static extern IntPtr xlAutoFree64(IntPtr handle);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "xlAutoFree12")]
        public static extern IntPtr xlAutoFree32(IntPtr handle);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "RtdCall")]
        public static extern IntPtr RtdCall64(IntPtr args);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "RtdCall")]
        public static extern IntPtr RtdCall32(IntPtr args);

        public static void RegisterFunction(Function function)
        {
            using (var ptr = new StructIntPtr<Function>(ref function))
            {
                if (Environment.Is64BitProcess)
                    xlAutoFree64(RegisterFunction64(ptr));
                else
                    xlAutoFree32(RegisterFunction32(ptr));
            }
        }

        public static void AsyncReturn(IntPtr handle, IntPtr result)
        {
            if (Environment.Is64BitProcess)
                xlAutoFree64(AsyncReturn64(handle, result));
            else
                xlAutoFree32(AsyncReturn32(handle, result));
        }

        public static void RtdCall(IntPtr args)
        {
            IntPtr result = IntPtr.Zero;
            if (Environment.Is64BitProcess)
                result = RtdCall64(args);
            else
                result = RtdCall32(args);
            //return Converter.ConvertOutgoing(result);
        }
    }
}
