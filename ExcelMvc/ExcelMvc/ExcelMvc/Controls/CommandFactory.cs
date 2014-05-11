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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelMvc.Extensions;
using ExcelMvc.Views;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Controls
{
    /// <summary>
    /// Creates command on a sheet
    /// </summary>
    internal static class CommandFactory
    {
        public static void Create(Worksheet sheet, View host, Dictionary<string, Command> commands)
        {
            var names = (from Comment item in sheet.Comments select item.Shape.Name).ToList();
            names.Sort();

            Create(sheet, host, (GroupObjects)sheet.GroupObjects(), names, commands);
            Create(sheet, host, (Buttons)sheet.Buttons(), names, commands);
            Create(sheet, host, (CheckBoxes)sheet.CheckBoxes(), names, commands);
            Create(sheet, host, (OptionButtons)sheet.OptionButtons(), names, commands);
            Create(sheet, host, (ListBoxes)sheet.ListBoxes(), names, commands);
            Create(sheet, host, (DropDowns)sheet.DropDowns(), names, commands);
            Create(sheet, host, (Spinners)sheet.Spinners(), names, commands);
            Create(sheet, host, sheet.Shapes, names, commands);
        }

        private static void Create(Worksheet sheet, View host, IEnumerable items, List<string> names, Dictionary<string, Command> commands)
        {
            foreach (var item in items)
            {
               var button = item as Button;
               if (Create(host, button, () => button.Name, () => button.OnAction, () => new CommandButton(host, button), commands, names))
                   continue;

               var cbox = item as CheckBox;
               if (Create(host, cbox, () => cbox.Name, () => cbox.OnAction, () => new CommandCheckBox(host, cbox), commands, names))
                   continue;

               var option = item as OptionButton;
               if (Create(host, option, () => option.Name, () => option.OnAction, () => new CommandOptionButton(host, option), commands, names))
                   continue;

               var lbox = item as ListBox;
               if (Create(host, lbox, () => lbox.Name, () => lbox.OnAction, () => new CommandListBox(host, lbox), commands, names))
                   continue;

               var dbox = item as DropDown;
               if (Create(host, dbox, () => dbox.Name, () => dbox.OnAction, () => new CommandDropDown(host, dbox), commands, names))
                   continue;

               var spin = item as Spinner;
               if (Create(host, spin, () => spin.Name, () => spin.OnAction, () => new CommandSpinner(host, spin), commands, names))
                   continue;

               if (Create(sheet, host, item as GroupObject, commands, names))
                   continue;

                if (Create(host, item as Shape, commands, names))
                {

                }
            }
        }

        private static bool Create(View host, object item, Func<string> getName, Func<string> getAction,
            Func<Command> createCmd, Dictionary<string, Command> commands, 
            List<string> names)
        {
            if (item == null)
                return false;

            var name = getName();
            var onAction = getAction();
            int idx;
            if ((idx = names.BinarySearch(name)) >= 0 || !IsCreateable(host, onAction))
                return true;

            ActionExtensions.Try(() =>
            {
                commands[name] = createCmd();
                names.Insert(~idx, name);
            });

            return true;
        }

        private static bool Create(Worksheet sheet, View host, GroupObject item, Dictionary<string, Command> commands, List<string> names)
        {
            if (item == null)
                return false;

            var name = item.Name;
            int idx;
            if ((idx = names.BinarySearch(name)) >= 0 || !IsCreateable(host, null))
                return true;

            ActionExtensions.Try(() =>
            {
                names.Insert(~idx, name);
                var shapes = (from Shape x in item.ShapeRange from Shape y in x.GroupItems select y).ToArray();
                item.Ungroup();
                Create(sheet, host, shapes, names, commands);
                sheet.Shapes.Range[(from Shape x in shapes select x.Name).ToArray()].Regroup();
            });
            return true;
        }

        private static bool Create(View host, Shape item, Dictionary<string, Command> commands, List<string> names)
        {
            if (item == null)
                return false;

            var name = item.Name;
            int idx;
            if ((idx = names.BinarySearch(name)) >= 0 || !IsCreateable(host, null))
                return true;

            ActionExtensions.Try(() =>
            {
                names.Insert(~idx, name);
                GroupShapes unused = null;
                ActionExtensions.Try(() => unused = item.GroupItems);
                if (unused == null)
                    commands[name] = new CommandShape(host, item);
            });

            return true;
        }

        private static bool IsCreateable(View host, string action)
        {
            return string.IsNullOrEmpty(action) || action == MacroNames.CommandActionName;
        }
    }
}
