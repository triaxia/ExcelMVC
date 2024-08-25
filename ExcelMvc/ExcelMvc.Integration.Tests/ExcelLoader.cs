using System;
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

        private static string AddInName
            => Path.Combine(Path.GetDirectoryName(typeof(BoolTests).Assembly.Location),
                 "ExcelMvc.Tests64.xll");

        public Application Application { get; }
        public AddIn AddIn { get; }
        public int ProcessId { get; }

        public ExcelLoader()
        {
            Application = new Application { Visible = false };
            var book = Application.Workbooks.Add();
            AddIn = Application.AddIns.Add(AddInName);
            AddIn.Installed = true;
            GetWindowThreadProcessId(Application.Hwnd, out var id);
            ProcessId = id;
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
