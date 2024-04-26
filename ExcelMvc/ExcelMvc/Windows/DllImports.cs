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
        internal static extern int SetTimer(int hwnd, int nIDEvent, int uElapse, IntPtr lpTimerFunc);

        [DllImport("user32")]
        internal static extern int KillTimer(IntPtr hwnd, int nIDEvent);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsCallback callback, IntPtr param);

        [DllImport("user32.dll")]
        internal static extern int GetClassNameW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);

        internal static Application FindExcel()
        {
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
            if (App.Instance.Underlying == null) return false;
            return IsFunctionWizardWindow(new IntPtr(App.Instance.Underlying.Hwnd));
        }

        internal static bool IsFunctionWizardWindow(IntPtr hWnd)
        {
            StringBuilder buffer = new StringBuilder(256);
            if (GetClassNameW(hWnd, buffer, buffer.Capacity) > 0
                 && buffer.ToString().StartsWith("bosa_sdm_XL", StringComparison.InvariantCultureIgnoreCase))
                return true;

            EnumChildWindows(hWnd, delegate (IntPtr hChild, IntPtr param)
            {
                if (IsFunctionWizardWindow(hChild))
                    return true;
                return false;
            }, IntPtr.Zero);
            return false;
        }
    }
}
