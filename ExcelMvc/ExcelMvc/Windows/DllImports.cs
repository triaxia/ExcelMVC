using ExcelMvc.Views;
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
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

        internal static uint MainNativeThreadId { get; set; }
        internal static Application FindExcel()
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
    }
}
