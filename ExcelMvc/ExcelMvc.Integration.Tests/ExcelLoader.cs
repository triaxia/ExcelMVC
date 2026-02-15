using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Integration.Tests
{
    public class ExcelLoader : IDisposable
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        
        private static string AddInName(bool is64)
            => Path.Combine(Path.GetDirectoryName(typeof(BoolTests).Assembly.Location),
                 is64 ? $"ExcelMvc.Tests64.xll" : $"ExcelMvc.Tests.xll");

        public Application Application { get; }
        public AddIn AddIn { get; }
        public int ProcessId { get; }

        [DllImport("kernel32.dll")]
        static extern bool IsWow64Process(IntPtr aProcessHandle, out bool lpSystemInfo);
        public static bool Is64BitProcess(IntPtr aProcessHandle)
        {
            if (!System.Environment.Is64BitOperatingSystem)
                return false;

            if (!IsWow64Process(aProcessHandle, out bool isWow64Process))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            return !isWow64Process;
        }

        public ExcelLoader()
        {
            Application = new Application { Visible = false };
            GetWindowThreadProcessId(Application.Hwnd, out var id);
            ProcessId = id;
            
            var is64 = Is64BitProcess(Process.GetProcessById(ProcessId).Handle);
            var addIn = AddInName(is64);

            var book = Application.Workbooks.Add();
            AddIn = Application.AddIns.Add(addIn);
            AddIn.Installed = true;
        }

        public void Dispose()
        {
            try
            {
                AddIn.Installed = false;
                Application.Quit();
            }
            catch { }
            finally
            {
                Process.GetProcessById(ProcessId)?.Kill();
            }
        }
    }
}
