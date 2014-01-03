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
using System.Collections;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Interop;
using ExcelMvc.Bindings;
using ExcelMvc.Controls;
using ExcelMvc.Views;
using Sample.Views;
using Binding = ExcelMvc.Bindings.Binding;
using Form = ExcelMvc.Views.Form;
using View = ExcelMvc.Views.View;

namespace Sample.Application.ViewModels
{
    internal class Forbes
    {
        private readonly Sheet _forbesSheet;
        private readonly Table _companyTable;
        private readonly Form _companyForm;
        private readonly Table _countryTable;
        private readonly Table _industryTable;

        public Forbes(View view)
        {
            view.HookBindingFailed(_view_BindFailed, true);

            _forbesSheet = (Sheet) view.Find(Binding.ViewType.Sheet, "Forbes");
            _forbesSheet.HookClicked(LoadAllClicked, "LoadForbes", true);
            _forbesSheet.HookClicked(ClearAllClicked, "ClearForbes", true);
            _forbesSheet.HookClicked(StartUpdateClicked, "StartUpdate", true);
            _forbesSheet.HookClicked(ShowColumnClicked, "ShowColumn", true);
            _forbesSheet.HookClicked(ShowDialogClicked, "ShowDialog", true);

            _companyTable = (Table)_forbesSheet.Find(Binding.ViewType.Table, "Company");
            _companyTable.SelectionChanged += _companyTable_SelectionChanged;
            _companyTable.ObjectChanged += _companyTable_ObjectChanged;
            _companyTable.Model = new CompanyList();

            _companyForm = (Form)_forbesSheet.Find(Binding.ViewType.Form, "Company");
            _companyForm.ObjectChanged += _companyForm_ObjectChanged;

            _countryTable = (Table) view.Find(Binding.ViewType.Table, "Country");
            _industryTable = (Table) view.Find(Binding.ViewType.Table, "Industry");
        }

        void _companyForm_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            // this is just for demo purpose, just to get the table to update, careful with 
            // recursive update
            (args.Items.First() as Company).RaiseChanged();
        }

        void _companyTable_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            var model = args.Items.Last();
            if (model == _companyForm.Model)
                ((Company) model).RaiseChanged();
        }

        void _companyTable_SelectionChanged(object sender, SelectionChangedArgs args)
        {
            _companyForm.Model = args.Items.Last();
        }

        private void _view_BindFailed(object sender, BindingFailedEventArgs args)
        {
            MessageBox.Show(args.Exception.Message, args.View.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        void LoadAllClicked(object sender, CommandEventArgs args)
        {
            var companyList =(CompanyList) _companyTable.Model;
            companyList.Load();
            RebindReferenceLists(companyList);
            companyList.RaiseChanged();

            _forbesSheet.FindCommand("LoadForbes").IsEnabled = false;
            _forbesSheet.FindCommand("ClearForbes").IsEnabled = true;
        }

        void ClearAllClicked(object sender, CommandEventArgs args)
        {
            var companyList = (CompanyList)_companyTable.Model;
            companyList.Unload();
            RebindReferenceLists(companyList);

            _forbesSheet.FindCommand("LoadForbes").IsEnabled = true;
            _forbesSheet.FindCommand("ClearForbes").IsEnabled = false;
        }

        private void RebindReferenceLists(CompanyList clist)
        {
            _countryTable.Model = clist.CountryList;
            _industryTable.Model = clist.IndustryList;
        }

        void StartUpdateClicked(object sender, CommandEventArgs args)
        {
            var cmd = (Command) sender;
            var update = !(bool) (cmd.Value ?? false);
            cmd.Value = update;
            cmd.Caption = update ? "Stop Update" : "Start Update";
            var companyList = (CompanyList)_companyTable.Model;
            companyList.Update(update);
        }

        private void ShowColumnClicked(object sender, CommandEventArgs args)
        {
            var visible =_companyTable.ToggleColumnVisibility("Industry");
            var cmd = (Command)sender;
            cmd.Caption = visible ? "Hide Industry" : "Show Industry";
        }

        private void ShowDialogClicked(object sender, CommandEventArgs args)
        {
            //var items = _companyTable.SelecteItems;
            //var v = new Forbes { Model = items };
            var v = new Forbes2000 { Model = (IEnumerable)_companyTable.Model };
            var interop = new WindowInteropHelper(v) {Owner = App.Instance.Root.Handle};
            v.ShowDialog();
        }
    }
}
