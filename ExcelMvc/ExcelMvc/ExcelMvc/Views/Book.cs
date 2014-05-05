/*
Copyright (c) 2013 Peter Gu or otherwise indicated by the license information contained within
the source files.

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
using System.Collections.Generic;
using System.Linq;
using ExcelMvc.Bindings;
using ExcelMvc.Controls;
using ExcelMvc.Extensions;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Views
{
    /// <summary>
    /// Represents a visual over an Excel workbook
    /// </summary>
    public class Book : View
    {
        /// <summary>
        /// The underlying Excel Workkbook
        /// </summary>
        public Workbook Underlying { get; private set; }

        private readonly Dictionary<Worksheet, Sheet> _sheets = new Dictionary<Worksheet, Sheet>();

        internal Book(View parent, Workbook book)
        {
            Parent = parent;
            Underlying = book;
        }

        public override void Dispose()
        {
            if (Underlying != null)
            {
                Underlying.SheetActivate -= Underlying_SheetActivate;
                Underlying.SheetDeactivate -= Underlying_SheetDeactivate;
            }
            foreach (var item in _sheets.Values)
                item.Dispose();
            _sheets.Clear();
        }

        /// <summary>
        /// Gets the child views
        /// </summary>
        public override IEnumerable<View> Children
        {
            get { return _sheets.Values.ToList() as IEnumerable<View>; }
        }

        /// <summary>
        /// Gets the book name
        /// </summary>
        public override string Name
        {
            get { return Underlying.Name; }
        }

        /// <summary>
        /// Gets the full book name
        /// </summary>
        public string FullName
        {
            get { return Underlying.FullName; }
        }

        /// <summary>
        /// Gets the book id, as defined by the Custom Document Propety named "ExcelMvc"
        /// </summary>
        public override string Id
        {
            get
            {
                var value = "";
                ActionExtensions.Try(() => value = ((Microsoft.Office.Core.DocumentProperties)Underlying.CustomDocumentProperties)[App.ExcelMvc].Value.ToString());
                return value;
            }
        }

        public override Binding.ViewType Type
        {
            get { return Binding.ViewType.Book; }
        }


        internal void Initialise()
        {
            Dispose();

            var bindings = Binding.Collect(Underlying);
            foreach (Worksheet item in Underlying.Worksheets)
            {
                var view = new Sheet(this, item);
                List<Binding> sheetBindings;
                bindings.TryGetValue(item, out sheetBindings);
                view.Initialise(sheetBindings);
                _sheets[item] = view;
            }
            Underlying.SheetActivate += Underlying_SheetActivate;
            Underlying.SheetDeactivate += Underlying_SheetDeactivate;
        }

        void Underlying_SheetActivate(object sh)
        {
            OnActivated(new ViewEventArgs(_sheets[(Worksheet)sh]));
        }

        void Underlying_SheetDeactivate(object sh)
        {
            OnDeactivated(new ViewEventArgs(_sheets[(Worksheet)sh]));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        internal Command FindCommand(Worksheet sheet, string name)
        {
            Command cmd = null;
            Sheet item;
            if (_sheets.TryGetValue(sheet, out item))
                cmd = item.FindCommand(name);
            return cmd;
        }
    }
}
