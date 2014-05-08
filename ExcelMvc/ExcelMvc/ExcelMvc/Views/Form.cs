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
using System.ComponentModel;
using System.Linq;
using ExcelMvc.Bindings;
using ExcelMvc.Runtime;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Views
{
    /// <summary>
    /// Represents a visual consists with scattered fields
    /// </summary>
    public class Form : BindingView
    {
        public override Binding.ViewType Type
        {
            get { return Binding.ViewType.Form; }
        }

        public override object Model
        {
            set
            {
                base.Model = value;
                HookModelEvents();
                UpdateView();
            }
        }

        /// <summary>
        /// Gets the selected bindings.
        /// </summary>
        public List<Binding> SelectedBindings { get; private set; }


        /// <summary>
        /// Initialises an instances of ExcelMvc.Views.Panel
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="bindings">Bindings for the view</param>
        internal Form(View parent, IEnumerable<Binding> bindings)
            : base(parent, bindings)
        {
            SelectedBindings = new List<Binding>();
        }

        private void HookViewEvents()
        {
            var sheet = (Sheet)Parent;
            sheet.Underlying.Change += Underlying_Change;
            sheet.Underlying.SelectionChange += Underlying_SelectionChange;
        }

        private void UnhookViewEvents()
        {
            var sheet = (Sheet)Parent;
            sheet.Underlying.Change -= Underlying_Change;
            sheet.Underlying.SelectionChange -= Underlying_SelectionChange;
        }

        void Underlying_Change(Range target)
        {
            UpdateObject(target);
        }

        private void Underlying_SelectionChange(Range target)
        {
            var count = SelectedBindings.Count;
            SelectedBindings.Clear();
            SelectedBindings.AddRange(Bindings.Where(binding => target.Application.Intersect(binding.Cell, target) != null));
            if (count != 0 || SelectedBindings.Count != 0)
                OnSelectionChanged(new [] { Model }, SelectedBindings);
        }

        void UpdateObject(Range target)
        {
            var toSource = Bindings.Where(x => (x.Mode == Binding.ModeType.TwoWay || x.Mode == Binding.ModeType.OneWayToSource));
            foreach (var binding in toSource)
                UpdateObject(binding, target);
        }

        private void UpdateObject(Binding binding, Range target)
        {
            ExcuteBinding(() =>
            {
                var range = binding.Cell;
                var changed = target.Application.Intersect(range, target);
                if (changed != null)
                {
                    var value = RangeConversion.MergeChangedValue(changed, range, ObjectBinding.GetPropertyValue(Model, binding));
                    if (value.Changed)
                    {
                        ObjectBinding.SetPropertyValue(Model, binding, value.Value);
                        OnObjectChanged(new[] { Model }, new[] { binding.Path });
                    }
                }
            });
        }

        private void UpdateView()
        {
            try
            {
                UnhookViewEvents();
                UpdateView("");
                BindValidationLists(1);
            }
            finally
            {
                HookViewEvents();
            }
        }

        private void UpdateView(string path)
        {
            ExcuteBinding(() =>
            {
                var  match = string.IsNullOrEmpty(path)  ? null : Bindings.FirstOrDefault(x => x.Path == path);
                if (match != null)
                {
                    UpdateView(match);
                }
                else
                {
                    foreach (var binding in Bindings)
                        UpdateView(binding);
                }
            });
        }

        private void UpdateView(Binding binding)
        {
            if (binding.Mode == Binding.ModeType.OneWayToSource)
                return;

            ExcuteBinding(() =>
            {
                var value = ObjectBinding.GetPropertyValue(Model, binding);
                RangeUpdator.Instance.Update(binding.Cell, 0, 1, 0, 1, value);
            });
        }

        private INotifyPropertyChanged _notifyPropertyChanged;
        private void HookModelEvents()
        {
            UnhookModelEvents();
            _notifyPropertyChanged = Model as INotifyPropertyChanged;
            if (_notifyPropertyChanged != null)
                _notifyPropertyChanged.PropertyChanged += notify_PropertyChanged;
        }

        private void UnhookModelEvents()
        {
            if (_notifyPropertyChanged != null)
                _notifyPropertyChanged.PropertyChanged -= notify_PropertyChanged;
        }

        void notify_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            UpdateView(e.PropertyName);
        }

        public override void Dispose()
        {
            base.Model = null;
            UnhookModelEvents();
            UnhookViewEvents();
        }
    }
}
