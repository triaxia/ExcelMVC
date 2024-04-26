using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelMvc.Windows
{
    internal class DllImports
    {
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

        internal static Application FindExcel()
        {
            var pid = Process.GetCurrentProcess().Id;
            IRunningObjectTable prot = null;
            IEnumMoniker pMonkEnum = null;
            try
            {
                DllImports.GetRunningObjectTable(0, out prot);
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
                        DllImports.GetWindowThreadProcessId(new IntPtr(excel.Hwnd), out uint excelpid);
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
    }
}
