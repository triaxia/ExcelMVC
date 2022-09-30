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
        public static void Register(ExcelFunction function)
        {
            var ptr = IntPtr.Zero;
            try
            {
                ptr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(ExcelFunction)));
                Marshal.StructureToPtr(function, ptr, false);
                if (Environment.Is64BitProcess)
                    RegisterFunction64(ptr);
                else
                    RegisterFunction32(ptr);
            }
            finally
            {
                if (ptr != IntPtr.Zero) Marshal.FreeHGlobal(ptr);
            }
        }
    }
}
