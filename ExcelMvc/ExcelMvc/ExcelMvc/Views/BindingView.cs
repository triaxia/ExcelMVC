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
using System.Reflection;
using ExcelMvc.Bindings;
using ExcelMvc.Controls;
using ExcelMvc.Extensions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Shape = Microsoft.Office.Interop.Excel.Shape;

namespace ExcelMvc.Views
{
    /// <summary>
    /// Defines an abstract interface for Views
    /// </summary>
    public abstract class BindingView : View
    {
        public override string Id
        {
            get { return Name; }
        }

        public override string Name
        {
            get { return Bindings.First().Name; }
        }

        /// <summary>
        /// Gets the bindings on the View
        /// </summary>
        public IEnumerable<Binding> Bindings { get; private set; }

        /// <summary>
        /// Occurs when a command is clicked
        /// </summary>
        public event ClickedHandler Clicked = delegate { };

        /// <summary>
        /// Initialises an instances of ExcelMvc.Views.Panel
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="bindings">Bindings for the view</param>
        protected BindingView(View parent, IEnumerable<Binding> bindings)
        {
            Bindings = (from o in bindings orderby o.Cell.Row, o.Cell.Column select o).ToList();
            Parent = parent;
        }

        /// <summary>
        /// Fires the Clicked event
        /// </summary>
        public void FireClicked(object sender, CommandEventArgs args)
        {
            Clicked(sender, args);
        }

        /// <summary>
        /// Unbinds validation lists
        /// </summary>
        /// <param name="rows">Number of rows to unbind</param>
        protected void UnbindValidationLists(int rows)
        {
            foreach (var binding in Bindings.Where(binding => !string.IsNullOrEmpty(binding.ValidationList)))
            {
                if (IsBoolValidationList(binding.ValidationList))
                    UnbindCheckBoxes(binding, rows);
                else
                    UnbindValidationLists(binding, rows);
            }
        }

        private void UnbindValidationLists(Binding binding, int rows)
        {
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                var column = lbinding.MakeRange(0, rows, 0, 1);
                column.Validation.Delete();
            });
        }

        private void UnbindCheckBoxes(Binding binding, int rows)
        {
            /*
            var worksheet = ((Sheet)Parent).Underlying;
            CheckBoxes boxes = worksheet.CheckBoxes();
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                for (var idx = 0; idx < rows; idx++)
                {
                    var cell = lbinding.MakeRange(idx, 1, 0, 1);
                    var name = "_ExcelMvc_" + cell.Address;
                    CheckBox box = null;
                    ActionExtensions.Try(() => box = boxes.Item(name));
                    if (box != null)
                        box.Delete();
                }
            });*/
        }

        /// <summary>
        /// Unbinds validation lists
        /// </summary>
        /// <param name="rows">Number of rows to unbind</param>
        protected void BindValidationLists(int rows)
        {
            foreach (var binding in Bindings.Where(binding => !string.IsNullOrEmpty(binding.ValidationList)))
            {
                if (IsBoolValidationList(binding.ValidationList))
                    BindCheckBoxes(binding, rows);
                else
                    BindValidationLists(binding, rows);
            }
        }

        private void BindValidationLists(Binding binding, int rows)
        {
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                var column = lbinding.MakeRange(0, rows, 0, 1);
                column.Validation.Delete();
                column.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                    XlFormatConditionOperator.xlBetween, MarkValidationListFormula(lbinding.ValidationList));
                column.Validation.IgnoreBlank = true;
                column.Validation.InCellDropdown = true;
                column.Validation.InputTitle = "";
                column.Validation.ErrorTitle = "";
                column.Validation.InputMessage = "";
                column.Validation.ErrorMessage = "";
                column.Validation.ShowInput = true;
                column.Validation.ShowError = true;
            });
        }

        private void BindCheckBoxes(Binding binding, int rows)
        {
            var worksheet = ((Sheet)Parent).Underlying;
            var boxes = worksheet.Shapes;
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                //for (var idx = 0; idx < rows; idx++)
                {
                    var cell = lbinding.MakeRange(0, 1, 0, 1);
                    Shape box = boxes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, cell.Left + 2, cell.Top + 2, 12, 12);
                    box.Fill.Visible = MsoTriState.msoFalse;
                    box.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                    box.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    box.TextFrame2.TextRange.Characters.Text = "X";
                    box.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = 192;
                    box.Line.Weight = 1;
                    cell.Select();
                   
                    var range = lbinding.MakeRange(0, 1000, 0, 1);
                    cell.AutoFill(range);
                    //range.FillDown();
                }
            });
        }

        private static bool IsBoolValidationList(string list)
        {
            return list.CompareOrdinalIgnoreCase("True/False") == 0;
        }

        private string MarkValidationListFormula(string list)
        {
            Range range;
            if (list.Contains("["))
            {
                range = ((Sheet)(Parent)).Underlying.Application.Range[list];
            }
            else if (list.Contains("!"))
            {
                var names = list.Split('!');
                range = ((Sheet)Parent.Parent.Find(Binding.ViewType.Sheet, names[0])).Underlying.Range[names[1]];
            }
            else
            {
                range = ((Sheet)Parent).Underlying.Range[list];
            }

            // exclude trailing blank rows
            var value = (object[,])range.Value;
            for (var idx = value.GetUpperBound(0); idx >= value.GetLowerBound(0); idx--)
            {
                if (value[idx, 1] == null)
                    continue;
                var rows = idx - value.GetLowerBound(0) + 1;
                range = range.Worksheet.Range[range.Cells[1, 1], range.Cells[rows, 1]];
                break;
            }

            var address = range.Address[Missing.Value, Missing.Value, XlReferenceStyle.xlA1, true, Missing.Value];
            return string.Format("={0}", address);
        }
    }
}
