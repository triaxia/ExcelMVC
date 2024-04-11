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
namespace ExcelMvc.Bindings
{
    using System;
    using System.Collections.Generic;
    using System.Windows.Data;

    using Extensions;
    using Microsoft.Office.Interop.Excel;
    using Runtime;
    using Views;
    using Range = Microsoft.Office.Interop.Excel.Range;
    internal class BindingCollector
    {

        public BindingCollector(Workbook book)
        {
            Book = book;
        }

        private string BindingPath
        {
            get;
            set;
        }

        private string StartCell
        {
            get;
            set;
        }

        private string EndCell
        {
            get;
            set;
        }

        private Dictionary<Worksheet, List<Binding>> Bindings
        {
            get;
            set;
        }

        private IValueConverter Converter
        {
            get;
            set;
        }

        private bool IsVisible
        {
            get;
            set;
        }

        private string ModeType
        {
            get;
            set;
        }

        private Range StartRange
        {
            get;
            set;
        }

        private Range EndRange
        {
            get;
            set;
        }

        private string Validation
        {
            get;
            set;
        }

        private string ViewName
        {
            get;
            set;
        }

        private string ViewType
        {
            get;
            set;
        }

        private Workbook Book
        {
            get;
            set;
        }

        public Dictionary<Worksheet, List<Binding>> Process()
        {
            Bindings = new Dictionary<Worksheet, List<Binding>>();

            foreach (Name nm in Book.Names)
                Collect(nm);

            foreach (Worksheet item in Book.Worksheets)
            {
                foreach (Name nm in item.Names)
                    Collect(nm);
            }

            return Bindings;
        }

        private static int IndexOfHeading(object[,] value, string heading)
        {
            var idx = value.GetLowerBound(0);
            for (var jdx = value.GetLowerBound(1); jdx <= value.GetUpperBound(1); jdx++)
            {
                var cell = value[idx, jdx] as string;
                if (cell != null && cell.EqualNoCase(heading))
                    return jdx;
            }

            return -1;
        }

        private static bool IsNotValidExcelMvcName(string[] parts)
        {
            return parts.Length != 3 || !parts[0].EqualNoCase("ExcelMvc");
        }

        private void Collect(Name nm)
        {
            var parts = nm.Name.Split('.');

            if (IsNotValidExcelMvcName(parts))
                return;

            CheckFirstPartOfName(parts);

            CheckSecondPartOfName(parts);

            var value = (object[,])nm.RefersToRange.Value;

            var indices = GetHeadingIndices(value);

            for (var idx = value.GetLowerBound(0) + 1; idx <= value.GetUpperBound(0); idx++)
            {
                StartCell = ((value[idx, indices.IndexOfStartCell] as string) ?? string.Empty).Trim();
                EndCell = indices.IndexOfEndCell >= 0 ? ((value[idx, indices.IndexOfEndCell] as string) ?? string.Empty).Trim() : string.Empty;

                BindingPath = ((value[idx, indices.IndexOfPath] as string) ?? string.Empty).Trim();
                if (StartCell == string.Empty || BindingPath == string.Empty)
                    continue;

                StartRange = GetRange(StartCell, nm);
                EndRange = string.IsNullOrWhiteSpace(EndCell) ? null : GetRange(EndCell, nm);

                ModeType = ((value[idx, indices.IndexOfMode] as string) ?? string.Empty).Trim();
                IsVisible = GetVisibility(value, indices, idx);

                Validation = GetValidation(value, indices, idx);

                Converter = GetConverter(value, indices, idx);

                var binding = CreateBinding();
                AddToBindings(binding);
            }
        }

        private void AddToBindings(Binding binding)
        {
            List<Binding> sheetBindings;
            var sheet = StartRange.Worksheet;
            if (!Bindings.TryGetValue(sheet, out sheetBindings))
                Bindings[sheet] = sheetBindings = new List<Binding>();
            sheetBindings.Add(binding);
        }

        private Binding CreateBinding()
        {
            var binding = new Binding
            {
                Name = ViewName,
                Type = (ViewType)Enum.Parse(typeof(ViewType), ViewType),
                Mode = (ModeType)Enum.Parse(typeof(ModeType), ModeType),
                StartCell = (Range)StartRange.Cells[1, 1],
                EndCell = EndRange == null ? null : (Range)EndRange.Cells[1, 1],
                Path = BindingPath,
                Visible = IsVisible,
                ValidationList = Validation,
                Converter = Converter
            };
            return binding;
        }

        private IValueConverter GetConverter(object[,] value, Indices indices, int idx)
        {
            if (indices.IndexOfConverter >= 0)
            {
                var typeName = value[idx, indices.IndexOfConverter] as string;
                if (!string.IsNullOrWhiteSpace(typeName))
                    return ObjectFactory<IValueConverter>.Find(typeName);
            }

            return null;
        }

        private string GetValidation(object[,] value, Indices indices, int idx)
        {
            if (indices.IndexOfValidation >= 0)
                return value[idx, indices.IndexOfValidation] as string;
            return null;
        }

        private Range GetRange(string cellAddress, Name nm)
        {
            Range range = null;
            ActionExtensions.Try(() =>
            {
                if (cellAddress.Contains("["))
                {
                    range = Book.Application.Range[cellAddress];
                }
                else if (cellAddress.Contains("!"))
                {
                    var names = cellAddress.Split('!');
                    range = (Book.Sheets[names[0]] as Worksheet).Range[names[1]];
                }
                else
                {
                    range = nm.RefersToRange.Worksheet.Range[cellAddress];
                }
            });
            if (range == null)
                throw new Exception(string.Format(Resource.ErrorNoDataCellRange, cellAddress));
            return range;
        }

        private bool GetVisibility(object[,] value, Indices indices, int idx)
        {
            if (indices.IndexOfVisibility >= 0)
            {
                var cell = (value[idx, indices.IndexOfVisibility] as string) ?? string.Empty;
                return cell.EqualNoCase("Visible") || cell.EqualNoCase("True");
            }

            return true;
        }

        private void CheckSecondPartOfName(string[] parts)
        {
            ViewName = parts[2] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(ViewName))
                throw new Exception(string.Format(Resource.ErrorNoViewNameFound, ViewName));
        }

        private void CheckFirstPartOfName(string[] parts)
        {
            ViewType = parts[1] ?? string.Empty;
            if (!ViewType.EqualNoCase("Form") && !ViewType.EqualNoCase("Table"))
                throw new Exception(string.Format(Resource.ErrorInvalidViewType, ViewType));
        }

        private Indices GetHeadingIndices(object[,] value)
        {
            Indices result;
            result.IndexOfStartCell = IndexOfHeading(value, "Start Cell");

            // for backward-compatibility, try "Data Cell"
            if (result.IndexOfStartCell == -1)
                result.IndexOfStartCell = IndexOfHeading(value, "Data Cell");
            if (result.IndexOfStartCell == -1)
                throw new Exception(Resource.ErrorNoBindingCellFound);

            result.IndexOfEndCell = IndexOfHeading(value, "End Cell");

            result.IndexOfPath = IndexOfHeading(value, "Binding Path");
            if (result.IndexOfPath == -1)
                throw new Exception(Resource.ErrorNoBindingPathFound);

            result.IndexOfMode = IndexOfHeading(value, "Binding Mode");
            if (result.IndexOfMode == -1)
                throw new Exception(Resource.ErrorNoBindingModeFound);

            result.IndexOfVisibility = IndexOfHeading(value, "Visibility");
            result.IndexOfValidation = IndexOfHeading(value, "Validation");
            result.IndexOfConverter = IndexOfHeading(value, "Converter");

            return result;
        }

        private struct Indices
        {

            public int IndexOfStartCell;
            public int IndexOfEndCell;
            public int IndexOfConverter;
            public int IndexOfMode;
            public int IndexOfPath;
            public int IndexOfValidation;
            public int IndexOfVisibility;
        }
    }
}
