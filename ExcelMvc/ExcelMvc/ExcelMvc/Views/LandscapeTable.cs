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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using ExcelMvc.Bindings;
using ExcelMvc.Extensions;
using ExcelMvc.Runtime;
using Binding = ExcelMvc.Bindings.Binding;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Views
{
    /// <summary>
    /// Represents a rectangular transposed visual with rows and columns
    /// </summary>
    public class LandscapeTable : BindingView
    {
        private IList<object> _itemsBound;
        private IEnumerable _enumerable;
        private INotifyPropertyChanged _notifyPropertyChanged;
        private INotifyCollectionChanged _notifyCollectionChanged;

        public override Binding.ViewType Type
        {
            get { return Binding.ViewType.LandscapeTable; }
        }

        /// <summary>
        /// Gets the selected objects.
        /// </summary>
        public List<object> SelectedItems { get; private set; }

        /// <summary>
        /// Gets the selected bindings.
        /// </summary>
        public List<Binding> SelectedBindings { get; private set; }

        /// <summary>
        /// Sets the underlying model
        /// </summary>
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
        /// Initialises an instances of ExcelMvc.Views.Panel
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="bindings">Bindings for the view</param>
        internal LandscapeTable(View parent, IEnumerable<Binding> bindings)
            : base(parent, bindings)
        {
            SelectedItems = new List<object>();
            SelectedBindings = new List<Binding>();
            SetRowVisibility();
        }

        private void HookViewEvents(bool isHook)
        {
            var sheet = (Sheet)Parent;
            if (isHook)
            {
                sheet.Underlying.SelectionChange += Underlying_SelectionChange;
                sheet.Underlying.Change += Underlying_Change;
            }
            else
            {
                sheet.Underlying.SelectionChange -= Underlying_SelectionChange;
                sheet.Underlying.Change -= Underlying_Change;
            }
        }

        void Underlying_Change(Range target)
        {
            RestoreColumnIds(target);
            UpdateRangeObjects(target);
        }

        void Underlying_SelectionChange(Range target)
        {
            SaveColumnIds(target);
            var rangeObjs = GetRangeObjects(target);
            SelectedItems.Clear();
            SelectedBindings.Clear();
            if (rangeObjs.Items != null)
            {
                SelectedItems.AddRange(rangeObjs.Items);
                SelectedBindings.AddRange(rangeObjs.Bindings);
                OnSelectionChanged(rangeObjs.Items, rangeObjs.Bindings);
            }
        }

        private void HookModelEvents()
        {
            _enumerable = Model as IEnumerable;
            if (_enumerable == null && Model != null)
                throw new Exception(string.Format(Resource.ErrorNoIEnumuerable, Model.GetType().FullName, Name));

            UnhookModelEvents();

            _notifyPropertyChanged = Model as INotifyPropertyChanged;
            if (_notifyPropertyChanged != null)
                _notifyPropertyChanged.PropertyChanged += _notifyPropertyChanged_PropertyChanged;

            _notifyCollectionChanged = Model as INotifyCollectionChanged;
            if (_notifyCollectionChanged != null)
                _notifyCollectionChanged.CollectionChanged += _notifyCollectionChanged_CollectionChanged;
        }

        private void UnhookModelEvents()
        {
            if (_notifyPropertyChanged != null)
                _notifyPropertyChanged.PropertyChanged -= _notifyPropertyChanged_PropertyChanged;
            if (_notifyCollectionChanged != null)
                _notifyCollectionChanged.CollectionChanged -= _notifyCollectionChanged_CollectionChanged;
        }

        private void HookItemsEvents(bool toHook)
        {
            if (_itemsBound == null)
                return;

            foreach (var item in _itemsBound)
            {
                var itemNotify = item as INotifyPropertyChanged;
                if (itemNotify != null)
                {
                    if (toHook)
                        itemNotify.PropertyChanged += itemNotify_PropertyChanged;
                    else
                        itemNotify.PropertyChanged -= itemNotify_PropertyChanged;
                }
            }
        }

        void itemNotify_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            UpdateCell(sender, e.PropertyName);
        }

        void _notifyCollectionChanged_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdateView();
        }

        void _notifyPropertyChanged_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
        }

        private void UpdateView()
        {
            try
            {
                HookItemsEvents(false);
                HookViewEvents(false);
                ExcuteBinding(UpdateViewEx);
            }
            finally
            {
                HookViewEvents(true);
                HookItemsEvents(true);
            }
        }

        private void UpdateViewEx()
        {
            Dictionary<Binding, List<object>> bindingValues = null;
            var colsBound = _itemsBound == null ? 0 : _itemsBound.Count;
            var toView = from binding in Bindings
                         where binding.Mode != Binding.ModeType.OneWayToSource
                         select binding;

            if (_enumerable != null)
            {
                bindingValues = new Dictionary<Binding, List<object>>();
                foreach (var binding in toView)
                    bindingValues[binding] = new List<object>();

                _itemsBound = _enumerable.ToList();

                foreach (var item in _itemsBound)
                    foreach (var binding in toView)
                        bindingValues[binding].Add(ObjectBinding.GetPropertyValue(item, binding));
            }

            var newCols = _itemsBound == null ? 0 : _itemsBound.Count;
            var groupBindings = GroupBindings(toView);
            if (colsBound != newCols)
            {
                ClearView(groupBindings, colsBound);
                if (colsBound > 0)
                    UnbindValidationLists(colsBound);
            }

            if (newCols > 0)
            {
                AssignColumnIds();
                UpdateView(groupBindings, bindingValues, newCols);
                BindValidationLists(newCols);
            }
        }

        private static List<List<Binding>> GroupBindings(IEnumerable<Binding> bindings)
        {
            var groups = new List<List<Binding>>();
            var ordered = (from x in bindings orderby x.Cell.Column select x).ToList();
            while (ordered.Count > 0)
            {
                int idx;
                for (idx = 1; idx < ordered.Count; idx++)
                {
                    if (ordered[idx].Cell.Row != ordered[idx - 1].Cell.Row + 1)
                        break;
                }
                groups.Add(ordered.Take(idx).ToList());
                ordered = ordered.Skip(idx).ToList();
            }
            return groups;
        }

        private static void UpdateView(IEnumerable<List<Binding>> groups, IDictionary<Binding, List<object>> bindingValues, int cols)
        {
            foreach (var group in groups)
                UpdateRange(group, bindingValues, cols);
        }

        private static void ClearView(IEnumerable<List<Binding>> groups, int cols)
        {
            foreach (var group in groups)
                ClearRange(group, cols);
        }

        private static void UpdateRange(List<Binding> bindings, IDictionary<Binding, List<object>> bindingValues, int cols)
        {
            var first = bindings[0];
            var rows = bindings.Count;
            var cells = new object[rows, cols];
            for (var idx = 0; idx < bindings.Count; idx++)
            {
                var values = bindingValues[bindings[idx]];
                for (var jdx = 0; jdx < cols; jdx++)
                    cells[idx, jdx] = values[jdx];
            }
            RangeUpdator.Instance.Update(first.Cell, 0, rows, 0, cols, cells);
        }

        private static void ClearRange(List<Binding> bindings, int cols)
        {
            var first = bindings[0];
            var rows = bindings.Count;
            var cells = new object[rows, cols];
            RangeUpdator.Instance.Update(first.Cell, 0, rows, 0, cols, cells);
        }

        void UpdateRangeObjects(Range target)
        {
            var from = target;
            while (target != null)
            {
                if (UpdateObjects(target))
                    break;

                Range dependents = null;
                var ltarget = target;
                ActionExtensions.Try(() => dependents = ltarget.DirectDependents);
                target = dependents;

                // break if circular referenced
                if (target != null && from.Application.Intersect(from, target) != null)
                    break;
            }
        }

        private bool UpdateObjects(Range target)
        {
            var updated = false;
            ExcuteBinding(() =>
            {
                var rangeObjs = GetRangeObjects(target);
                if (rangeObjs.Items != null)
                {
                    updated = true;
                    if (UpdateObjects(rangeObjs) > 0)
                        OnObjectChanged(rangeObjs.Items, null);
                }
            });
            return updated;
        }

        private int UpdateObjects(RangeObjects rangeItems)
        {
            var first = Bindings.First();
            var target = rangeItems.Intersection;
            var colOffset = target.Column - first.Cell.Column;
            var toSource = Bindings.Skip(target.Row - first.Cell.Row).Take(target.Rows.Count)
                .Where(x => (x.Mode == Binding.ModeType.TwoWay || x.Mode == Binding.ModeType.OneWayToSource)).ToList();
            var updated = 0;
            foreach (var model in rangeItems.Items)
            {
                updated += toSource.Count(binding => UpdateObject(binding, colOffset, model, rangeItems.Intersection));
                colOffset++;
            }
            return updated;
        }

        private static bool UpdateObject(Binding binding, int colOffset, object model, Range target)
        {
            var range = binding.MakeRange(0, 1, colOffset, 1);
            var changed = range.Application.Intersect(range, target);
            var value = RangeConversion.MergeChangedValue(changed, range, ObjectBinding.GetPropertyValue(model, binding));
            if (value.Changed)
                ObjectBinding.SetPropertyValue(model, binding, value.Value);
            return value.Changed;
        }

        private struct RangeObjects
        {
            public IEnumerable<object> Items;
            public IEnumerable<Binding> Bindings;
            public Range Intersection;
        }

        private RangeObjects GetRangeObjects(Range target)
        {
            RangeObjects result;
            result.Items = null;
            result.Bindings = null;
            result.Intersection = null;
            if (_itemsBound == null || _itemsBound.Count == 0)
                return result;

            var first = Bindings.First();
            var whole = first.MakeRange(0, Bindings.Count(), 0, _itemsBound.Count);
            result.Intersection = target.Application.Intersect(whole, target);
            if (result.Intersection == null)
                return result;

            var items = new List<object>();
            foreach (Range col in result.Intersection.Columns)
            {
                var cell = (Range)col.Worksheet.Cells[first.Cell.Row, col.Column];
                items.Add(_itemsBound[int.Parse(cell.ID)]);
            }
            result.Items = items;
            result.Bindings = Bindings.Skip(result.Intersection.Row - first.Cell.Row).Take(result.Intersection.Rows.Count).ToList();

            return result;
        }

        private void UpdateCell(object model, string propertyName)
        {
            var objectId = _itemsBound.IndexOf(model);
            if (objectId < 0)
                return;

            var match = Bindings.FirstOrDefault(x => x.Path == propertyName);
            if (match != null)
            {
                UpdateCell(match, model, objectId);
            }
            else
            {
                foreach (var binding in Bindings)
                    UpdateCell(binding, model, objectId);
            }
        }

        private void UpdateCell(Binding binding, object model, int objectId)
        {
            ExcuteBinding(() =>
            {
                var value = ObjectBinding.GetPropertyValue(model, binding);
                RangeUpdator.Instance.Update(binding.Cell, 0, 1, Bindings.First().Cell, _itemsBound.Count,
                    objectId.ToString(CultureInfo.InvariantCulture), 1, value);
            });
        }

        public override void Dispose()
        {
            HookViewEvents(false);
            HookItemsEvents(false);

            base.Model = null;
            UnhookModelEvents();
        }

        private void AssignColumnIds()
        {
            var first = Bindings.First();
            var row = first.MakeRange(0, 1, 0, _itemsBound.Count);
            Parent.ExecuteProtected(() =>
            {
                for (var idx = 1; idx <= _itemsBound.Count; idx++)
                    ((Range)row.Cells[1, idx]).ID = (idx - 1).ToString(CultureInfo.InvariantCulture);
            });
        }

        private readonly List<string> _colIds = new List<string>();
        private void SaveColumnIds(Range seletion)
        {
            _colIds.Clear();
            var first = Bindings.First();
            var row = first.MakeRange(0, 1, 0, _itemsBound.Count);
            var section = seletion.Application.Intersect(row, seletion);
            if (section != null)
            {
                foreach (Range col in section.Columns)
                    _colIds.Add(((Range)col.Cells[1, 1]).ID);
            }
            RestoreColumnIds(seletion);
        }

        private void RestoreColumnIds(Range seletion)
        {
            if (_colIds.Count == 0)
                return;

            var first = Bindings.First();
            var row = first.MakeRange(0, 1, 0, _itemsBound.Count);
            var section = seletion.Application.Intersect(row, seletion);
            if (section != null)
            {
                Parent.ExecuteProtected(() =>
                {
                    for (var idx = 1; idx <= section.Columns.Count; idx++)
                        ((Range)section.Cells[1, idx]).ID = _colIds[idx - 1];
                });
            }
        }

        /// <summary>
        /// Sets visibiliy for all rows
        /// </summary>
        public void SetRowVisibility()
        {
            Parent.ExecuteProtected(() =>
            {
                foreach (var binding in Bindings)
                    binding.Cell.EntireRow.Hidden = !binding.Visible;
            });
        }

        /// <summary>
        /// Toggles the visibility of a row
        /// </summary>
        /// <param name="path">Binding path of the row</param>
        /// <returns>true if visible, false otherwise</returns>
        public bool ToggleRowVisibility(string path)
        {
            var binding = Bindings.FirstOrDefault(x => x.Path == path);
            if (binding != null)
            {
                binding.Visible = !binding.Visible;
                Parent.ExecuteProtected(() => binding.Cell.EntireRow.Hidden = !binding.Visible);
            }
            return binding != null && binding.Visible;
        }
    }
}
