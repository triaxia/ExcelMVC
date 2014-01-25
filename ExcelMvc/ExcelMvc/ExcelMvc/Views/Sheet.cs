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
using System.Linq;
using ExcelMvc.Bindings;
using Microsoft.Office.Interop.Excel;
using ExcelMvc.Controls;
using ExcelMvc.Extensions;

namespace ExcelMvc.Views
{
    /// <summary>
    /// Represents a visual over an Excel worksheet
    /// </summary>
    public class Sheet : View
    {
        private readonly Dictionary<string, Command> _commands =
            new Dictionary<string, Command>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, Form> _forms =
            new Dictionary<string, Form>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, Table> _tables =
            new Dictionary<string, Table>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Occurs when a command is clicked
        /// </summary>
        public event ClickedHandler Clicked = delegate { };

        /// <summary>
        /// The underlying Excel sheet
        /// </summary>
        public Worksheet Underlying { get; private set; }

        /// <summary>
        /// Gets the Commands on the sheet
        /// </summary>
        public override IEnumerable<Command> Commands
        {
            get { return _commands.Values.ToList(); }
        }

        /// <summary>
        /// Gets the child views
        /// </summary>
        public override IEnumerable<View> Children
        {
            get
            {
                var list = _forms.Values.Select(x => (View)x).ToList();
                list.AddRange(_tables.Values.ToList());
                return list;
            }
        }

        public override string Id
        {
            get { return Name; }
        }

        /// <summary>
        /// Gets the workspace name
        /// </summary>
        public override string Name
        {
            get { return Underlying.Name; }
        }

        public override Binding.ViewType Type
        {
            get { return Binding.ViewType.Sheet; }
        }

        /// <summary>
        /// Initiaalises an instance of ExcelMvc.Views.Workspace
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="sheet">The underlying Excel Worksheet</param>
        internal Sheet(View parent, Worksheet sheet)
        {
            Parent = parent;
            Underlying = sheet;
        }

        internal void Initialise(IEnumerable<Binding> bindings)
        {
            Dispose();

            if (bindings != null)
                CreateViews(bindings);
            CreateCommands();
        }

        public override void Dispose()
        {
            foreach (var cmd in _commands.Values)
                cmd.Dispose();
            _commands.Clear();

            foreach (var form in _forms.Values)
                form.Dispose();
            _forms.Clear();

            foreach (var table in _tables.Values)
                table.Dispose();
            _tables.Clear();
        }

        private void CreateViews(IEnumerable<Binding> bindings)
        {
            var names = bindings.Where(x => x.Type == Binding.ViewType.Form).Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (var item in names)
            {
                var name = item;
                var fields = bindings.Where(x => x.Type == Binding.ViewType.Form && x.Name.CompareOrdinalIgnoreCase(name) == 0);
                var form = new Form(this, fields);
                var args = new ViewEventArgs(form);
                OnOpening(args);
                if (!args.IsCancelled)
                {
                    _forms[name] = form;
                    OnOpened(args);
                }
            }

            names = bindings.Where(x => x.Type == Binding.ViewType.Table).Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (var item in names)
            {
                var name = item;
                var columns = bindings.Where(x => x.Type == Binding.ViewType.Table && x.Name.CompareOrdinalIgnoreCase(name) == 0);
                var table = new Table(this, columns);
                var args = new ViewEventArgs(table);
                OnOpening(args);
                if (!args.IsCancelled)
                {
                    _tables[name] = table;
                    OnOpened(new ViewEventArgs(table));
                }
            }
        }

        private void CreateCommands()
        {
            var sheet = Underlying;
            var names = (from Comment item in sheet.Comments select item.Shape.Name).ToList();
            names.Sort();

            CreateCommands(sheet.GroupObjects(), names);
            CreateCommands(sheet.Buttons(), names);
            CreateCommands(sheet.CheckBoxes(), names);
            CreateCommands(sheet.OptionButtons(), names);
            CreateCommands(sheet.ListBoxes(), names);
            CreateCommands(sheet.DropDowns(), names);
            CreateCommands(sheet.Spinners(), names);
            CreateCommands(sheet.Shapes, names);

            foreach (var cmd in _commands.Values)
                cmd.Clicked += cmd_Clicked;
        }

        private void CreateCommands(dynamic items, List<string> names) 
        {
            foreach (dynamic item in items)
            {
                int idx;
                if ((idx = names.BinarySearch(item.Name)) >= 0 || !IsCreateable(item.OnAction))
                    continue;

                Command cmd = null;
                var litem = item;
                ActionExtensions.Try(() =>
                {
                    if (litem is Button)
                        cmd = new CommandButton(this, litem);
                    else if (litem is CheckBox)
                        cmd = new CommandCheckBox(this, litem);
                    else if (litem is OptionButton)
                        cmd = new CommandOptionButton(this, litem);
                    else if (litem is ListBox)
                        cmd = new CommandListBox(this, litem);
                    else if (litem is DropDown)
                        cmd = new CommandDropDown(this, litem);
                    else if (litem is Spinner)
                        cmd = new CommandSpinner(this, litem);
                    else if (litem is GroupObject)
                    {
                        names.Insert(~idx, litem.Name);
                        var gitem = (GroupObject)litem;
                        var shapes = (from Shape x in gitem.ShapeRange from Shape y in x.GroupItems select y).ToArray();
                        gitem.Ungroup();
                        CreateCommands(shapes, names);
                        Underlying.Shapes.Range[(from Shape x in shapes select x.Name).ToArray()].Regroup();
                    }
                    else if (litem is Shape)
                    {
                        var shape = litem as Shape;
                        GroupShapes unused = null;
                        ActionExtensions.Try(() => unused = shape.GroupItems);
                        if (unused == null)
                            cmd = new CommandShape(this, litem);
                    }
                });

                if (cmd != null)
                {
                    _commands[item.Name] = cmd;
                    names.Insert(~idx, item.Name);
                }
            }
        }

        private static bool IsCreateable(string action)
        {
            return string.IsNullOrEmpty(action) || action == MacroNames.CommandActionName;
        }

        void cmd_Clicked(object sender, CommandEventArgs args)
        {
            Clicked(sender, args);
            if (args.Handled)
                return;

            var views = _forms.Values.Select(x => x as BindingView).ToList();
            views.AddRange(_tables.Values.Select(x => x as BindingView));
            foreach (var view in views)
            {
                view.FireClicked(sender, args);
                if (args.Handled)
                    return;
            }
        }
    }
}
