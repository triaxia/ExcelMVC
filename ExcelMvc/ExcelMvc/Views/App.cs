/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany (2013)

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

namespace ExcelMvc.Views
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Data;
    using Bindings;
    using Controls;
    using ExcelMvc.Functions;
    using ExcelMvc.Windows;
    using Function.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using Runtime;

    using Application = Microsoft.Office.Interop.Excel.Application;

    /// <summary>
    /// Represents a visual over the Excel Application
    /// </summary>
    public class App : View
    {
        private static readonly Dictionary<Workbook, Book> Books = new Dictionary<Workbook, Book>();

        static App()
        {
            Instance = new App();
        }

        /// <summary>
        /// Disallow instance creation
        /// </summary>
        private App()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        public static string ExcelMvc
        {
            get { return "ExcelMvc"; }
        }

        /// <summary>
        /// Gets the singleton instance of ExcelMvc.Views.Books 
        /// </summary>
        public static App Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// 
        /// </summary>
        public override IEnumerable<View> Children
        {
            get { return Books.Values.ToArray(); }
        }

        /// <summary>
        /// 
        /// </summary>
        public override IEnumerable<Command> Commands
        {
            get { return new Command[] { }; }
        }

        /// <summary>
        /// 
        /// </summary>
        public override string Id
        {
            get { return ExcelMvc; }
        }

        /// <summary>
        /// 
        /// </summary>
        public override string Name
        {
            get { return ExcelMvc; }
        }

        /// <summary>
        /// Excel Main Window
        /// </summary>
        public Root MainWindow
        {
            get;
            private set;
        }

        /// <summary>
        /// 
        /// </summary>
        public override ViewType Type
        {
            get { return ViewType.App; }
        }

        /// <summary>
        /// The underlying Excel.Application instance
        /// </summary>
        public Application Underlying
        {
            get;
            private set;
        }

        /// <summary>
        /// Disposes resources
        /// </summary>
        public override void Dispose()
        {
            Detach();
            Underlying = null;
        }

        /// <summary>
        /// Collects bindings and rebinds the view
        /// </summary>
        /// <param name="recursive"></param>
        public override void Rebind(bool recursive)
        {
            foreach (var view in Children)
                view.Rebind(recursive);
        }

        /// <summary>
        /// Attaches the Excel Application instance to this instance
        /// </summary>
        internal void Attach(object app)
        {
            if (Underlying == null)
                DoAttach(app);
            else
                AsyncActions.PostMacro(a => DoAttach(a), app);
        }

        void DoAttach(object state)
        {
            Try(() =>
            {
                FunctionHost.Instance = new ExcelFunctionHost();

                Detach();
                Underlying = (state as Application) ?? DllImports.FindExcel();
                if (Underlying == null) throw new Exception(Resource.ErrorExcelAppFound);

                AsyncActions.Initialise();
                RaisePosted($"AsyncActions.Initialise() done");

                ObjectFactory<IFunctionAddIn>.CreateAll(ObjectFactory<IFunctionAddIn>.GetCreatableTypes
                    , ObjectFactory<IFunctionAddIn>.SelectAllAssembly);
                RaisePosted($"ObjectFactory<IFunctionAddIn>.CreateAll({ObjectFactory<IFunctionAddIn>.Instances.Count})");

                ObjectFactory<ISession>.CreateAll(ObjectFactory<ISession>.GetCreatableTypes
                    , ObjectFactory<ISession>.SelectAllAssembly);
                RaisePosted($"ObjectFactory<ISession>.CreateAll({ObjectFactory<ISession>.Instances.Count})");

                ObjectFactory<IValueConverter>.CreateAll(ObjectFactory<IValueConverter>.GetCreatableTypes
                    , ObjectFactory<IValueConverter>.SelectAllAssembly);
                RaisePosted($"ObjectFactory<IValueConverter>.CreateAll({ObjectFactory<IValueConverter>.Instances.Count})");

                var functions = FunctionDiscovery.RegisterFunctions();
                RaisePosted($"FunctionDiscovery.RegisterFunctions({functions.FunctionCount})");

                Underlying.WorkbookOpen += OpenBook;
                Underlying.WorkbookBeforeClose += ClosingBook;
                Underlying.WorkbookActivate += Activate;
                Underlying.WorkbookDeactivate += Deactivate;

                MainWindow = new Root(Underlying.Hwnd);
                MainWindow.Destroyed += MainWindow_Destroyed;

                foreach (Workbook item in Underlying.Workbooks)
                {
                    var view = new Book(this, item);
                    var args = new ViewEventArgs(view);
                    OnOpening(args);
                    if (args.IsAccepted)
                    {
                        view.Initialise();
                        Books[item] = view;
                        ExecuteBinding(() => OnOpened(new ViewEventArgs(view)));
                    }
                }
            });
        }

        /// <summary>
        /// Detaches Excel from this instance
        /// </summary>
        internal void Detach()
        {
            Try(() =>
            {
                if (Underlying != null)
                {
                    Underlying.WorkbookOpen -= OpenBook;
                    Underlying.WorkbookBeforeClose -= ClosingBook;
                    Underlying.WorkbookActivate -= Activate;
                    Underlying.WorkbookDeactivate -= Deactivate;

                    // Underlying is needed when Attach is invoked from the surface of the sheet
                    // Underlying = null;
                }

                MainWindow = null;

                foreach (var book in Books.Values)
                    book.Dispose();
                Books.Clear();

                ObjectFactory<ISession>.DeleteAll(x => x.Dispose());
                ObjectFactory<IValueConverter>.DeleteAll(x => { });
                ObjectFactory<IFunctionAddIn>.DeleteAll(x => { });
            });
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
            Try(() =>
            {
                if (Underlying == null)
                    return;
                var caller = CommandFactory.RemovePrefix(Underlying.Caller as string);
                var cmd = FindCommand((Worksheet)Underlying.ActiveSheet, caller);
                if (cmd != null && cmd.IsEnabled)
                    cmd.FireClicked();
            });
        }

        private void Activate(Workbook book)
        {
            Purge();
            if (Books.ContainsKey(book))
                Try(() => OnActivated(new ViewEventArgs(Books[book])));
        }

        private void ClosingBook(Workbook book, ref bool cancel)
        {
            var toCancel = cancel;
            Try(() =>
            {
                if (Books.TryGetValue(book, out Book view))
                {
                    var args = new ViewEventArgs(view);
                    OnClosing(args);
                }
            });
            cancel = toCancel;
        }

        private void Deactivate(Workbook book)
        {
            Purge();
            if (Books.ContainsKey(book))
                Try(() => OnDeactivated(new ViewEventArgs(Books[book])));
        }

        private void OpenBook(Workbook book)
        {
            Try(() =>
            {
                Purge();

                var isCreated = Books.TryGetValue(book, out Book view);
                if (isCreated)
                    return;
                view = new Book(this, book);
                var args = new ViewEventArgs(view);
                OnOpening(args);
                if (args.IsAccepted)
                {
                    view.Initialise();
                    Books[book] = view;
                    OnOpened(new ViewEventArgs(view));
                }
            });
        }

        private void Purge()
        {
            Try(() =>
            {
                var books = (from object obj in Underlying.Workbooks select (Workbook)obj).ToList();
                foreach (var key in Books.Keys.ToArray())
                {
                    if (books.Any(x => ReferenceEquals(x, key)))
                        continue;
                    var view = Books[key];
                    Books.Remove(key);
                    OnClosed(new ViewEventArgs(view));
                    view.Dispose();
                }
            });
        }

        private void MainWindow_Destroyed(object sender, EventArgs args)
        {
            Try(() => OnDestroyed(this));
        }

        private void Try(System.Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                OnBindingFailed(new BindingFailedEventArgs(this, ex));
                RaiseFailed(ex);
            }
        }
    }
}
