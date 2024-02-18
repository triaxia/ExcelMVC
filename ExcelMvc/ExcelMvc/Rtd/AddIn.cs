using System;
using System.Runtime.InteropServices;
using static ExcelMvc.Rtd.RtdServerFactory;
namespace ExcelMvc.Rtd
{
    using HRESULT = Int32;
    using IID = Guid;
    using CLSID = Guid;

    public static unsafe class AddIn
    {
        public static string ModuleFileName { get; private set; }
        internal delegate HRESULT fn_dll_get_class_object(CLSID rclsid, IID riid, out IntPtr ppunk);
        [StructLayout(LayoutKind.Sequential)]
        public struct AddInHead
        {
            public IntPtr ModuleFileName;
            public IntPtr pDllGetClassObject;
        }

        public static void OnAttach(IntPtr head)
        {
            AddInHead* pAddInHead = (AddInHead*)head;
            ModuleFileName = Marshal.PtrToStringAuto(pAddInHead->ModuleFileName);
            fn_dll_get_class_object fnDllGetClassObject = (fn_dll_get_class_object)DllGetClassObject;
            GCHandle.Alloc(fnDllGetClassObject);
            pAddInHead->pDllGetClassObject = Marshal.GetFunctionPointerForDelegate(fnDllGetClassObject);
        }
    }
}
