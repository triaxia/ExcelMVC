#region Header
/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Developer:         Wolfgang Stamm, Germany

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor,
Boston, MA 02110-1301 USA.
*/
#endregion Header

namespace ExcelMvc.Views
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Windows.Data;

    using ExcelMvc.Controls;
    using ExcelMvc.Runtime;

    using Microsoft.Office.Interop.Excel;

    using Application = Microsoft.Office.Interop.Excel.Application;

    using Binding = ExcelMvc.Bindings.Binding;

    /// <summary>
    /// Represents a visual over the Excel Application
    /// </summary>
    public class App : View
    {
        #region Fields

        private static readonly Dictionary<Workbook, Book> Books = new Dictionary<Workbook, Book>();

        #endregion Fields

        #region Constructors

        static App()
        {
            Instance  = new App();
        }

        /// <summary>
        /// Disallow instance creation
        /// </summary>
        private App()
        {
        }

        #endregion Constructors

        #region Properties

        public static string ExcelMvc
        {
            get { return "ExcelMvc"; }
        }

        /// <summary>
        /// Gets the singleton instance of ExcelMvc.Views.Books 
        /// </summary>
        public static App Instance
        {
            get; private set;
        }

        public override IEnumerable<View> Children
        {
            get { return Books.Values.ToArray(); }
        }

        public override IEnumerable<Command> Commands
        {
            get { return new Command[] { }; }
        }

        public override string Id
        {
            get { return ExcelMvc; }
        }

        public override string Name
        {
            get { return ExcelMvc; }
        }

        /// <summary>
        /// Excel Main Window
        /// </summary>
        public Root Root
        {
            get; private set;
        }

        public override Binding.ViewType Type
        {
            get { return Binding.ViewType.App; }
        }

        /// <summary>
        /// The underlying Excel.Application instance
        /// </summary>
        public Application Underlying
        {
            get; private set;
        }

        #endregion Properties

        #region Methods

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        public override void Dispose()
        {
            Underlying = null;
        }

        /// <summary>
        /// Attaches the Excel Application instance to this instance
        /// </summary>
        internal void Attach(object app)
        {
            Detach();
            ObjectFactory<ISession>.CreateAll();
            ObjectFactory<IValueConverter>.CreateAll();

            Underlying = (app as Application) ?? Find();
            Underlying.WorkbookOpen += OpenBook;
            Underlying.WorkbookBeforeClose += CloseingBook;
            Underlying.WorkbookActivate += Activate;
            Underlying.WorkbookDeactivate += Deactivate;

            Root = new Root(Underlying.Hwnd);
            Root.Destroyed += MainWindow_Destroyed;

            foreach (Workbook item in Underlying.Workbooks)
            {
                var book = new Book(this, item);
                var args = new ViewEventArgs(book);
                OnOpening(args);
                if (!args.IsCancelled)
                {
                    book.Initialise();
                    OnOpened(args);
                    Books[item] = book;
                }
            }
        }

        /// <summary>
        /// Detaches Excel from this instance
        /// </summary>
        internal void Detach()
        {
            if (Underlying != null)
            {
                Underlying.WorkbookOpen -= OpenBook;
                Underlying.WorkbookBeforeClose -= CloseingBook;
                Underlying.WorkbookActivate -= Activate;
                Underlying.WorkbookDeactivate -= Deactivate;
                Underlying = null;
            }

            Root = null;

            foreach (var space in Books.Values)
                space.Dispose();
            Books.Clear();

            ObjectFactory<ISession>.DeleteAll(x => x.Dispose());
            ObjectFactory<IValueConverter>.DeleteAll(x => { });
        }

        /// <summary>
        /// Finds the command by a command name
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="name">Command name</param>
        /// <returns>null or the command found</returns>
        internal Command FindCommand(Worksheet sheet, string name)
        {
            Command cmd = null;
            foreach (var book in Books.Values)
            {
                cmd = book.FindCommand(sheet, name);
                if (cmd != null)
                    break;
            }

            return cmd;
        }

        /// <summary>
        /// Fires the clicked event
        /// </summary>
        internal void FireClicked()
        {
            if (Underlying == null)
                return;

            var caller = CommandFactory.RemovePrefix(Underlying.Caller as string);
            var cmd = FindCommand((Worksheet)Underlying.ActiveSheet, caller);
            if (cmd != null && cmd.IsEnabled)
                cmd.FireClicked();
        }

        private static Application Find()
        {
            Application excel = null;
            int hWnd = Process.GetCurrentProcess().MainWindowHandle.ToInt32();
            IRunningObjectTable prot = null;
            IEnumMoniker pMonkEnum = null;
            try
            {
                GetRunningObjectTable(0, out prot);
                prot.EnumRunning(out pMonkEnum);
                var pmon = new IMoniker[1];
                var fetched = IntPtr.Zero;
                while (excel == null && pMonkEnum.Next(1, pmon, fetched) == 0)
                {
                    object result;
                    prot.GetObject(pmon[0], out result);
                    var book = result as Workbook;
                    if (book != null && book.Application.Hwnd == hWnd)
                        excel = book.Application;
                }
            }
            finally
            {
                if (prot != null)
                    Marshal.ReleaseComObject(prot);
                if (pMonkEnum != null)
                    Marshal.ReleaseComObject(pMonkEnum);
            }

            return excel;
        }

        private void Activate(Workbook book)
        {
            Purge();
            OnActivated(new ViewEventArgs(Books[book]));
        }

        private void CloseingBook(Workbook book, ref bool cancel)
        {
            Book space;
            if (Books.TryGetValue(book, out space))
            {
                var args = new ViewEventArgs(space);
                OnClosing(args);
                cancel = cancel | args.IsCancelled;
            }
        }

        private void Deactivate(Workbook book)
        {
            if (Books.Count < 1)
                return;
            OnDeactivated(new ViewEventArgs(Books[book]));
        }

        private void OpenBook(Workbook book)
        {
            Book space;
            var created = Books.TryGetValue(book, out space);
            if (!created)
            {
                space = new Book(this, book);
                var args = new ViewEventArgs(space);
                OnOpening(args);
                if (!args.IsCancelled)
                {
                    space.Initialise();
                    Books[book] = space;
                    OnOpened(args);
                }
            }
        }

        private void Purge()
        {
            var books = (from object obj in Underlying.Workbooks select (Workbook)obj).ToList();
            foreach (var key in Books.Keys.ToArray())
            {
                if (!books.Any(x => ReferenceEquals(x, key)))
                {
                    var space = Books[key];
                    Books.Remove(key);
                    OnClosed(new ViewEventArgs(space));
                    space.Dispose();
                }
            }
        }

        private void MainWindow_Destroyed(object sender, EventArgs args)
        {
            OnDestroyed(this);
        }

        #endregion Methods
    }
}