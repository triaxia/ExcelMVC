using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Runtime
{
    [StructLayout(LayoutKind.Sequential)]
    public struct AddInHead
    {
        public IntPtr ModuleFileName;
        public IntPtr pDllGetClassObject;
    }
}
