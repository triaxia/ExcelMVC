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
    /// Represents a rectangular visual with rows and columns
    /// </summary>
    public class Table : BindingView
    {
        private IList<object> _itemsBound;
        private IEnumerable _enumerable;
        private INotifyPropertyChanged _notifyPropertyChanged;
        private INotifyCollectionChanged _notifyCollectionChanged;
        private TableOrientation _orientation;

        public override Binding.ViewType Type
        {
            get { return Binding.ViewType.Table; }
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

        public enum TableOrientation
        {
            Portrait,
            Landscape
        }

        /// <summary>
        /// Initialises an instances of ExcelMvc.Views.Panel
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="bindings">Bindings for the view</param>
        internal Table(View parent, IEnumerable<Binding> bindings, TableOrientation orientation)
            : base(parent, bindings)
        {
            _orientation = orientation;
            SelectedItems = new List<object>();
            SelectedBindings = new List<Binding>();
            SetCategoryVisibility();
        }

        private void HookViewEvents(bool isHook)
        {
            var sheet = (Sheet) Parent;
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
            RestoreCategoryIds(target);
            UpdateRangeObjects(target);
        }

        void Underlying_SelectionChange(Range target)
        {
            SaveCategoryIds(target);
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
            var numberItemsBound = _itemsBound == null ? 0  : _itemsBound.Count;
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

            var newItems = _itemsBound == null ? 0 : _itemsBound.Count;
            var groupBindings = GroupBindings(toView, _orientation);
            if (numberItemsBound != newItems)
            {
                ClearView(groupBindings, numberItemsBound, _orientation);
                if (numberItemsBound > 0)
                    UnbindValidationLists(numberItemsBound, _orientation);
            }

            if (newItems > 0)
            {
                AssignItemIds();
                UpdateView(groupBindings, bindingValues, newItems, _orientation);
                BindValidationLists(newItems, _orientation);
            }
        }

        private static List<List<Binding>> GroupBindings(IEnumerable<Binding> bindings, TableOrientation orientation)
        {
            var groups = new List<List<Binding>>();
            var ordered = (from x in bindings orderby orientation == TableOrientation.Portrait ? x.Cell.Column : x.Cell.Row select x).ToList();
            while (ordered.Count > 0)
            {
                int idx;
                for (idx = 1; idx < ordered.Count; idx++)
                {
                    var currentBindingCell = orientation == TableOrientation.Portrait ? ordered[idx].Cell.Column : ordered[idx].Cell.Row;
                    var successorOfPreviousBindingCell = orientation == TableOrientation.Portrait ? ordered[idx - 1].Cell.Column + 1 : ordered[idx - 1].Cell.Row + 1;
                    if (currentBindingCell != successorOfPreviousBindingCell)
                        break;
                }
                groups.Add(ordered.Take(idx).ToList());
                ordered = ordered.Skip(idx).ToList();
            }
            return groups;
        }

        private static void UpdateView(IEnumerable<List<Binding>> groups, IDictionary<Binding, List<object>> bindingValues, int numberItems, TableOrientation orientation)
        {
            foreach (var group in groups)
                UpdateRange(group, bindingValues, numberItems, orientation);
        }

        private static void ClearView(IEnumerable<List<Binding>> groups, int numberItems, TableOrientation orientation)
        {
            foreach (var group in groups)
                ClearRange(group, numberItems, orientation);
        }

        private static void UpdateRange(List<Binding> bindings, IDictionary<Binding, List<object>> bindingValues, int numberItems, TableOrientation orientation)
        {
            var first = bindings[0];
            var numberCategories = bindings.Count;
            var numberRows = orientation == TableOrientation.Portrait ? numberItems : numberCategories;
            var numberCols = orientation == TableOrientation.Portrait ? numberCategories : numberItems;
            var cells = new object[numberRows, numberCols];
            for (var categoryIndex = 0; categoryIndex < bindings.Count; categoryIndex++)
            {
                var values = bindingValues[bindings[categoryIndex]];
                for (var itemIndex = 0; itemIndex < numberItems; itemIndex++)
                {
                    var idx = orientation == TableOrientation.Portrait ? itemIndex : categoryIndex;
                    var jdx = orientation == TableOrientation.Portrait ? categoryIndex : itemIndex;
                    cells[idx, jdx] = values[itemIndex];
                }
            }
            RangeUpdator.Instance.Update(first.Cell, 0, numberRows, 0, numberCols, cells);
        }

        private static void ClearRange(List<Binding> bindings, int numberItems, TableOrientation orientation)
        {
            var first = bindings[0];
            var numberCategories = bindings.Count;
            var numberRows = orientation == TableOrientation.Portrait ? numberItems : numberCategories;
            var numberCols = orientation == TableOrientation.Portrait ? numberCategories : numberItems;
            var cells = new object[numberRows, numberCols];
            RangeUpdator.Instance.Update(first.Cell, 0, numberRows, 0, numberCols, cells);
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
            var itemOffset = _orientation == TableOrientation.Portrait ? target.Row - first.Cell.Row : target.Column - first.Cell.Column;
            var skipCategories = _orientation == TableOrientation.Portrait ? target.Column - first.Cell.Column : target.Row - first.Cell.Row;
            var takeCategories = _orientation == TableOrientation.Portrait ? target.Columns.Count : target.Rows.Count;
            var toSource = Bindings.Skip(skipCategories).Take(takeCategories)
                .Where(x => (x.Mode == Binding.ModeType.TwoWay || x.Mode == Binding.ModeType.OneWayToSource)).ToList();
            var updated = 0;
            foreach (var model in rangeItems.Items)
            {
                updated += toSource.Count(binding => UpdateObject(binding, itemOffset, model, rangeItems.Intersection, _orientation));
                itemOffset++;
            }
            return updated;
        }

        private static bool UpdateObject(Binding binding, int itemOffset, object model, Range target, TableOrientation orientation)
        {
            var range = orientation == TableOrientation.Portrait ? binding.MakeRange(itemOffset, 1, 0, 1) : binding.MakeRange(0, 1, itemOffset, 1);
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
            var whole = GetWholeRange(first);
            result.Intersection = target.Application.Intersect(whole, target);
            if (result.Intersection == null)
                return result;

            var items = new List<object>();
            foreach (Range item in _orientation == TableOrientation.Portrait ? result.Intersection.Rows : result.Intersection.Columns)
            {
                var idx = _orientation == TableOrientation.Portrait ? item.Row : first.Cell.Row;
                var jdx = _orientation == TableOrientation.Portrait ? first.Cell.Column : item.Column;
                var cell = (Range)item.Worksheet.Cells[idx, jdx];
                items.Add(_itemsBound[int.Parse(cell.ID)]);
            }
            result.Items = items;
            var skipItems = _orientation == TableOrientation.Portrait ? result.Intersection.Column - first.Cell.Column : result.Intersection.Row - first.Cell.Row;
            var takeItems = _orientation == TableOrientation.Portrait ? result.Intersection.Columns.Count : result.Intersection.Rows.Count;
            result.Bindings = Bindings.Skip(skipItems).Take(takeItems).ToList();

            return result;
        }

        private Range GetWholeRange(Binding binding)
        {
            switch (_orientation)
            {
                case TableOrientation.Portrait:
                    return binding.MakeRange(0, _itemsBound.Count, 0, Bindings.Count());
                case TableOrientation.Landscape:
                    return binding.MakeRange(0, Bindings.Count(), 0, _itemsBound.Count);
            }
            return null;
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
                switch (_orientation)
                {
                    case TableOrientation.Portrait:
                        RangeUpdator.Instance.Update(binding.Cell, Bindings.First().Cell, _itemsBound.Count,
                    objectId.ToString(CultureInfo.InvariantCulture), 1, 0, 1, value);
                        break;
                    case TableOrientation.Landscape:
                        RangeUpdator.Instance.Update(binding.Cell, 0, 1, Bindings.First().Cell, _itemsBound.Count,
                    objectId.ToString(CultureInfo.InvariantCulture), 1, value);
                        break;
                }
                
            });
        }

        public override void Dispose()
        {
            HookViewEvents(false);
            HookItemsEvents(false);

            base.Model = null;
            UnhookModelEvents();
        }

        private void AssignItemIds()
        {
            var first = Bindings.First();
            var categoryRange = GetCategoryRange(first);
            Parent.ExecuteProtected(() =>
            {
                for (var itemIndex = 1; itemIndex <= _itemsBound.Count; itemIndex++)
                {
                    var idx = _orientation == TableOrientation.Portrait ? itemIndex : 1;
                    var jdx = _orientation == TableOrientation.Portrait ? 1 : itemIndex;
                    ((Range)categoryRange.Cells[idx, jdx]).ID = (itemIndex - 1).ToString(CultureInfo.InvariantCulture);
                }
            });
        }

        private readonly List<string> _categoryIds  = new List<string>();
        private void SaveCategoryIds(Range selection)
        {
            _categoryIds.Clear();
            var first = Bindings.First();
            var categoryRange = GetCategoryRange(first);
            var section = selection.Application.Intersect(categoryRange, selection);
            if (section != null)
            {
                foreach (Range item in _orientation == TableOrientation.Portrait ? section.Rows : section.Columns)
                    _categoryIds.Add(((Range)item.Cells[1, 1]).ID);
            }
            RestoreCategoryIds(selection);
        }

        private Range GetCategoryRange(Binding binding)
        {
            switch (_orientation)
            {
                case TableOrientation.Portrait:
                    return binding.MakeRange(0, _itemsBound.Count, 0, 1);
                case TableOrientation.Landscape:
                    return binding.MakeRange(0, 1, 0, _itemsBound.Count);
            }
            return null;
        }

        private void RestoreCategoryIds(Range selection)
        {
            if (_categoryIds.Count == 0)
                return;

            var first = Bindings.First();
            var categoryRange = GetCategoryRange(first);
            var section = selection.Application.Intersect(categoryRange, selection);
            if (section != null)
            {
                Parent.ExecuteProtected(() =>
                {
                    for (var itemIndex = 1; itemIndex <= (_orientation == TableOrientation.Portrait ? section.Rows.Count : section.Columns.Count); itemIndex++)
                    {
                        var idx = _orientation == TableOrientation.Portrait ? itemIndex : 1;
                        var jdx = _orientation == TableOrientation.Portrait ? 1 : itemIndex;
                        ((Range)section.Cells[idx, jdx]).ID = _categoryIds[itemIndex - 1];
                    }
                });
            }
        }

        /// <summary>
        /// Sets visibiliy for all columns (portrait table) or for all rows (landscape table)
        /// </summary>
        public void SetCategoryVisibility()
        {
            Parent.ExecuteProtected(() =>
            {
                foreach (var binding in Bindings)
                    SetCategoryVisibility(binding);
            });
        }

        /// <summary>
        /// Sets visibiliy for a single column (portrait table) or for a single row (landscape table)
        /// </summary>
        private void SetCategoryVisibility(Binding binding)
        {
            switch (_orientation)
            {
                case TableOrientation.Portrait:
                    binding.Cell.EntireColumn.Hidden = !binding.Visible;
                    break;
                case TableOrientation.Landscape:
                    binding.Cell.EntireRow.Hidden = !binding.Visible;
                    break;
            }
        }

        /// <summary>
        /// Toggles the visibility of a column (portrait table) or row (landscape table)
        /// </summary>
        /// <param name="path">Binding path of the column or row</param>
        /// <returns>true if visible, false otherwise</returns>
        public bool ToggleCategoryVisibility(string path)
        {
            var binding = Bindings.FirstOrDefault(x => x.Path == path);
            if (binding != null)
            {
                binding.Visible = !binding.Visible;
                switch (_orientation)
                {
                    case TableOrientation.Portrait:
                        Parent.ExecuteProtected(() => binding.Cell.EntireColumn.Hidden = !binding.Visible);
                        break;
                    case TableOrientation.Landscape:
                        Parent.ExecuteProtected(() => binding.Cell.EntireRow.Hidden = !binding.Visible);
                        break;
                }
            }
            return binding != null && binding.Visible;
        }
    }
}
