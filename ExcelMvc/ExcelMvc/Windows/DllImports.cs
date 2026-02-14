using ExcelMvc.Views;
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ExcelMvc.Windows
{
    internal class DllImports
    {
        internal delegate bool EnumWindowsCallback(IntPtr hwnd, IntPtr param);

        [DllImport("ole32.dll")]
        internal static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        internal static extern int PostMessage(IntPtr hwnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern uint RegisterWindowMessage(string lpProcName);

        [DllImport("user32")]
        internal static extern int SetTimer(IntPtr hwnd, int nIDEvent, int uElapse, IntPtr lpTimerFunc);

        [DllImport("user32")]
        internal static extern int KillTimer(IntPtr hwnd, int nIDEvent);

        [DllImport("user32.dll")]
        internal static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsCallback callback, IntPtr param);

        [DllImport("user32.dll")]
        internal static extern bool EnumThreadWindows(uint dwThreadId, EnumWindowsCallback callback, IntPtr param);

        [DllImport("user32.dll")]
        internal static extern int GetClassNameW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);

        [DllImport("Kernel32")]
        internal static extern uint GetCurrentThreadId();
        
        [DllImport("Oleacc.dll")]
        internal static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, byte[] riid, ref IntPtr ptr /*ppUnk*/ );

        private const uint OBJID_NATIVEEOM = 0xFFFFFFF0;
        private static readonly byte[] IID_IDispatchBytes = new Guid("{00020400-0000-0000-C000-000000000046}").ToByteArray();
        private static readonly CultureInfo EnUsCulture = new CultureInfo(1033);

        internal static uint MainNativeThreadId { get; set; }

        internal static Application FindExcel()
        {
            return FindExcelFromWindows() ?? FindExcelFromRunningObjectTable();
        }

        internal static Application FindExcelFromWindows()
        {
            MainNativeThreadId = GetCurrentThreadId();
            Application app = null;
            EnumThreadWindows(MainNativeThreadId, delegate (IntPtr hWndEnum, IntPtr param)
            {
                if (IsXlMainWindow(hWndEnum))
                {
                    app = GetApplicationFromWindow(hWndEnum);
                    return app == null;
                }
                return true;
            }, IntPtr.Zero);
            return app;
        }

        internal static Application FindExcelFromRunningObjectTable()
        {
            MainNativeThreadId = GetCurrentThreadId();
            var pid = Process.GetCurrentProcess().Id;
            IRunningObjectTable prot = null;
            IEnumMoniker pMonkEnum = null;
            try
            {
                _ = GetRunningObjectTable(0, out prot);
                prot.EnumRunning(out pMonkEnum);
                var pmon = new IMoniker[1];
                var fetched = IntPtr.Zero;
                while (pMonkEnum.Next(1, pmon, fetched) == 0)
                {
                    prot.GetObject(pmon[0], out object result);
                    var excel = result as Application;
                    if (excel == null) excel = (result as Workbook)?.Application;
                    if (excel != null)
                    {
                        _ = GetWindowThreadProcessId(new IntPtr(excel.Hwnd), out uint excelpid);
                        if (pid == excelpid)
                            return excel;
                    }
                    Marshal.ReleaseComObject(result);
                }
            }
            finally
            {
                if (prot != null)
                    Marshal.ReleaseComObject(prot);
                if (pMonkEnum != null)
                    Marshal.ReleaseComObject(pMonkEnum);
            }

            return null;
        }

        internal static bool IsInFunctionWizard()
        {
            if (App.Instance.Underlying == null)
                return false;

            var buffer = new StringBuilder(256);
            var result = false;
            EnumThreadWindows(MainNativeThreadId, delegate (IntPtr hWndEnum, IntPtr param)
            {
                if (IsFunctionWizardWindow(hWndEnum))
                {
                    result = true;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        internal static bool IsFunctionWizardWindow(IntPtr hWnd)
        {
            var buffer = new StringBuilder(256);
            return GetClassNameW(hWnd, buffer, buffer.Capacity) > 0
                && buffer.ToString().StartsWith("bosa_sdm_XL", StringComparison.InvariantCultureIgnoreCase)
                && IsReallyFunctionWizardWindow(hWnd);
        }

        internal static bool IsReallyFunctionWizardWindow(IntPtr hWnd)
        {
            // Below is inspired by ExcelDna.Integration.ExcelDnaUtil.IsFunctionWizardWindow.
            // Well until a better way is found!
            var editBoxCount = 0;
            var scrollbarCount = 0;
            EnumChildWindows(hWnd, delegate (IntPtr hChild, IntPtr param)
            {
                var buffer = new StringBuilder(256);
                if (GetClassNameW(hChild, buffer, buffer.Capacity) == 0)
                    return true;

                var name = buffer.ToString();
                if (name == "EDTBX") editBoxCount++;
                if (name == "ScrollBar") scrollbarCount++;

                return true;
            }, IntPtr.Zero);

            return editBoxCount == 5 && scrollbarCount == 1;
        }

        internal static bool IsXlMainWindow(IntPtr hWnd)
        {
            var buffer = new StringBuilder(256);
            GetClassNameW(hWnd, buffer, buffer.Capacity);
            return buffer.ToString() == "XLMAIN";
        }

        private static Application GetApplicationFromWindow(IntPtr hWndMain)
        {
            Application app = null;
            var clsName = new StringBuilder(256);
            EnumChildWindows(hWndMain, delegate (IntPtr hWndEnum, IntPtr para)
            {
                GetClassNameW(hWndEnum, clsName, clsName.Capacity);
                if (clsName.ToString() != "EXCEL7")
                    // not a workbook, continue
                    return true;
                IntPtr pUnk = IntPtr.Zero;
                int hr = AccessibleObjectFromWindow(hWndEnum, OBJID_NATIVEEOM, IID_IDispatchBytes, ref pUnk);
                if (hr != 0) 
                    return true;
                object obj = Marshal.GetObjectForIUnknown(pUnk);
                Marshal.Release(pUnk);
                if (HasProperty(obj, "Application"))
                {
                    app = (Application)obj.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, obj, null, EnUsCulture);
                }
                /*
                else if (HasProperty(obj,"Workbook"))
                {
                    var workbook = obj.GetType().InvokeMember("Workbook", BindingFlags.GetProperty, null, obj, null, EnUsCulture);
                    app = (Application)workbook.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, workbook, null, EnUsCulture);

                }*/
                return app == null;
            }, IntPtr.Zero);
            return app;
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00020400-0000-0000-C000-000000000046")]
        interface IDispatch
        {
            [PreserveSig]
            int GetTypeInfoCount(out int count);

            [PreserveSig]
            int GetTypeInfo
            (
                [MarshalAs(UnmanagedType.U4)] int iTInfo,
                [MarshalAs(UnmanagedType.U4)] int lcid,
                out System.Runtime.InteropServices.ComTypes.ITypeInfo typeInfo
            );

            [PreserveSig]
            int GetIDsOfNames
            (
                ref Guid riid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)]
                      string[] rgsNames,
                uint cNames,
                int lcid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4, SizeParamIndex = 2)] int[] reDispId
            );

            [PreserveSig]
            int Invoke
            (
                int dispIdMember,
                ref Guid riid,
                uint lcid,
                ushort wFlags,
                ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
                out object pVarResult,
                ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
                out UInt32 pArgErr

            );
        }

        public static bool HasProperty(object dispatchObject, string name)
        {
            const int LcidUsEnglish = 0x0409;
            string[] names = new string[1];
            int[] ids = new int[1];
            const int S_OK = 0;

            var dispObj = dispatchObject as IDispatch;
            if ( dispObj == null)
                return false; 
            names[0] = name;
            int hr = dispObj.GetIDsOfNames(Guid.Empty, names, 1, LcidUsEnglish, ids);
            return hr == S_OK;
        }
    }
}
