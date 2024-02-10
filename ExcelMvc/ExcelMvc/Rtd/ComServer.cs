using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Rtd
{
    using HRESULT = Int32;
    using IID = Guid;
    using CLSID = Guid;
    internal delegate HRESULT fn_dll_get_class_object(CLSID rclsid, IID riid, out IntPtr ppunk);

    [StructLayout(LayoutKind.Sequential)]
    public struct AddInHead
    {
        public IntPtr ModuleFileName;
        public IntPtr pDllGetClassObject;
    }

    public static unsafe class ComServer
    {
        public const HRESULT S_OK = 0;
        public const HRESULT S_FALSE = 1;
        public const HRESULT CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
        public const HRESULT CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
        public const HRESULT E_ACCESSDENIED = unchecked((int)0x80070005);
        public const HRESULT E_INVALIDARG = unchecked((int)0x80070057);
        public const HRESULT E_NOINTERFACE = unchecked((int)0x80004002);
        public const HRESULT E_UNEXPECTED = unchecked((int)0x8000FFFF);

        public static string ModuleFileName { get; private set; }

        public static HRESULT DllGetClassObject(CLSID clsid, IID iid, out IntPtr ppunk)
        {
            HRESULT result = S_OK;
            ppunk = IntPtr.Zero;
            return result;
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
