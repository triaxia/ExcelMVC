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
        public static extern IntPtr AsyncReturn64(uint index, IntPtr function);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "AsyncReturn")]
        public static extern IntPtr AsyncReturn32(uint index, IntPtr function);

        public static void RegisterFunction(ExcelFunction function)
        {
            using (var ptr = new StructIntPtr<ExcelFunction>(ref function))
            {
                if (Environment.Is64BitProcess)
                    RegisterFunction64(ptr);
                else
                    RegisterFunction32(ptr);
            }
        }
        public static void AsyncReturn(uint index, IntPtr result)
        {
            if (Environment.Is64BitProcess)
                AsyncReturn64(index, result);
            else
                AsyncReturn32(index, result);
        }
    }
}
