using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Integration.Tests
{
    public class ExcelLoader : IDisposable
    {
        private static string AddInName
            => Path.Combine(Path.GetDirectoryName(typeof(FunctionTests).Assembly.Location),
                 "ExcelMvc.Tests64.xll");

        public Application Application { get; private set; }
        public AddIn AddIn { get; private set; }
       

        public ExcelLoader() 
        {
            Application = new Application { Visible = false };
            var book = Application.Workbooks.Add();
            AddIn = Application.AddIns.Add(AddInName);
            AddIn.Installed = true;
        }

        public void Dispose()
        {
            AddIn.Installed = false;
            Application.Quit();
            foreach (var excel in Process.GetProcessesByName("Excel"))
                excel.Kill();
        }
    }
}
