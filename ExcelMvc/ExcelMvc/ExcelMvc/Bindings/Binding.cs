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

using System;
using System.Collections.Generic;
using System.Windows.Data;
using ExcelMvc.Extensions;
using ExcelMvc.Runtime;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Bindings
{
    /// <summary>
    /// Represents either a form field binding or a table column binding between 
    /// the View (Excel) and its view model
    /// </summary>
    public class Binding
    {
        /// <summary>
        /// View types
        /// </summary>
        public enum ViewType
        {
            /// <summary>
            /// None
            /// </summary>
            None,

            /// <summary>
            /// Form
            /// </summary>
            Form,

            /// <summary>
            /// Table
            /// </summary>
            Table,

            /// <summary>
            /// Sheet
            /// </summary>
            Sheet,

            /// <summary>
            /// Book
            /// </summary>
            Book,

            /// <summary>
            /// App
            /// </summary>
            App
        }

        /// <summary>
        /// Gets and sets the View Type
        /// </summary>
        public ViewType Type { get; set; }

        /// <summary>
        /// Binding mode types
        /// </summary>
        public enum ModeType
        {
            OneWay,
            OneWayToSource,
            TwoWay
        }

        /// <summary>
        /// Gets and sets the mode Type
        /// </summary>
        public ModeType Mode { get; set; }

        /// <summary>
        /// View name
        /// </summary>
        internal string Name { get; set; }

        /// <summary>
        /// Property path
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Visual cell
        /// </summary>
        internal Range Cell { get; set; }

        /// <summary>
        /// Visible
        /// </summary>
        public bool Visible { get; set; }

        /// <summary>
        /// Validation list address
        /// </summary>
        public string ValidationList { get; set; }

        /// <summary>
        /// Value converter
        /// </summary>
        public IValueConverter Converter { get; set; }

        /// <summary>
        /// Initialises an instance of ExcelMvc.Binding
        /// </summary>
        public Binding()
        {
            Type = ViewType.None;
            Mode = ModeType.OneWay;
            Name = "";
            Path = "";
            Visible = true;
        }

        /// <summary>
        /// Makes a range from the binding cell
        /// </summary>
        /// <param name="rowOffset">Start row offset</param>
        /// <param name="rows">Rows to extend from the binding Cell</param>
        /// <param name="columnOffset">Start column offset</param>
        /// <param name="cols">Columns to extend from the binding Cell</param>
        /// <returns>Column range</returns>
        public Range MakeRange(int rowOffset, int rows, int columnOffset, int cols)
        {
            var start = Cell.Worksheet.Cells[Cell.Row + rowOffset, Cell.Column + columnOffset];
            var end = Cell.Worksheet.Cells[Cell.Row + rowOffset + rows - 1, Cell.Column + +columnOffset + cols - 1];
            return Cell.Worksheet.Range[start, end];
        }

        internal static Dictionary<Worksheet, List<Binding>> Collect(Workbook book)
        {
            var bindings = new Dictionary<Worksheet, List<Binding>>();

            foreach (Name nm in book.Names)
                Collect(nm, book, bindings);

            foreach (Worksheet item in book.Worksheets)
            {
                foreach (Name nm in item.Names)
                    Collect(nm, book, bindings);
            }
            return bindings;
        }

        private static void Collect(Name nm, Workbook book, Dictionary<Worksheet, List<Binding>> bindings)
        {
            var parts = nm.Name.Split('.');

            if (parts.Length != 3 || parts[0].CompareOrdinalIgnoreCase("ExcelMvc") != 0)
                return;

            var viewType = parts[1] ?? "";
            if (viewType.CompareOrdinalIgnoreCase("Form") != 0
                && viewType.CompareOrdinalIgnoreCase("Table") != 0)
                throw new Exception(string.Format(Resource.ErrorInvalidViewType, viewType));

            var viewName = parts[2] ?? "";
            if (string.IsNullOrEmpty(viewName))
                throw new Exception(string.Format(Resource.ErrorNoViewNameFound, viewName));

            object[,] value = (object[,])nm.RefersToRange.Value;
            var indexOfCell = IndexOfHeading(value, "Data Cell");
            if (indexOfCell == -1)
                throw new Exception(Resource.ErrorNoBindingCellFound);

            var indexOfPath = IndexOfHeading(value, "Binding Path");
            if (indexOfPath == -1)
                throw new Exception(Resource.ErrorNoBindingPathFound);

            var indexOfMode = IndexOfHeading(value, "Binding Mode");
            if (indexOfMode == -1)
                throw new Exception(Resource.ErrorNoBindingModeFound);

            var indexOfVisibility = IndexOfHeading(value, "Visibility");
            var indexOfValidation = IndexOfHeading(value, "Validation");
            var indexOfConverter = IndexOfHeading(value, "Converter");

            for (var idx = value.GetLowerBound(0) + 1; idx <= value.GetUpperBound(0); idx++)
            {
                var dataCell = ((value[idx, indexOfCell] as string) ?? "").Trim();
                var bindingPath = ((value[idx, indexOfPath] as string) ?? "").Trim();
                if (dataCell == "" || bindingPath == "")
                    continue;

                Range range = null;
                ActionExtensions.Try(() =>
                {
                    if (dataCell.Contains("["))
                    {
                        range = book.Application.Range[dataCell];
                    }
                    else if (dataCell.Contains("!"))
                    {
                        var names = dataCell.Split('!');
                        range = (book.Sheets[names[0]] as Worksheet).Range[names[1]];
                    }
                    else
                    {
                        range = nm.RefersToRange.Worksheet.Range[dataCell];
                    }
                });
                if (range == null)
                    throw new Exception(string.Format(Resource.ErrorNoDataCellRange, dataCell));

                var modeType = ((value[idx, indexOfMode] as string) ?? "").Trim();
                var visible = true;
                if (indexOfVisibility >= 0)
                {
                    var cell = (value[idx, indexOfVisibility] as string) ?? "";
                    visible = cell.CompareOrdinalIgnoreCase("Visible") == 0 || cell.CompareOrdinalIgnoreCase("True") == 0;
                }
                string validation = null;
                if (indexOfValidation >= 0)
                    validation = (value[idx, indexOfValidation] as string);

                IValueConverter converter = null;
                if (indexOfConverter >= 0)
                {
                    var typeName = value[idx, indexOfConverter] as string;
                    if (!string.IsNullOrEmpty(typeName))
                        converter = ObjectFactory<IValueConverter>.Find(typeName);
                }

                var binding = new Binding
                {
                    Name = viewName,
                    Type = (ViewType)Enum.Parse(typeof(ViewType), viewType),
                    Mode = (ModeType)Enum.Parse(typeof(ModeType), modeType),
                    Cell = (Range)range.Cells[1, 1],
                    Path = bindingPath,
                    Visible = visible,
                    ValidationList = validation,
                    Converter = converter
                };

                List<Binding> sheetBindings;
                if (!bindings.TryGetValue(range.Worksheet, out sheetBindings))
                    bindings[range.Worksheet] = sheetBindings = new List<Binding>();
                sheetBindings.Add(binding);
            }
        }

        private static int IndexOfHeading(object[,] value, string heading)
        {
            var idx = value.GetLowerBound(0);
            for (var jdx = value.GetLowerBound(1); jdx <= value.GetUpperBound(1); jdx++)
            {
                var cell = value[idx, jdx] as string;
                if (cell != null && cell.CompareOrdinalIgnoreCase(heading) == 0)
                    return jdx;
            }
            return -1;
        }
    }
}
