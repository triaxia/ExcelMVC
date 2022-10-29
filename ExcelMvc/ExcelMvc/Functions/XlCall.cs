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

        public static void RegisterFunction(Function function)
        {
            using (var ptr = new StructIntPtr<Function>(ref function))
            {
                if (Environment.Is64BitProcess)
                    RegisterFunction64(ptr);
                else
                    RegisterFunction32(ptr);
            }
        }
        public static void AsyncReturn(IntPtr handle, IntPtr result)
        {
            if (Environment.Is64BitProcess)
                AsyncReturn64(handle, result);
            else
                AsyncReturn32(handle, result);
        }
    }
}
