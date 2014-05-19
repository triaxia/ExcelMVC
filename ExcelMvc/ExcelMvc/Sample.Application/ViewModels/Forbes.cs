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
        private Sheet ForbesSheet { get; set; }
        private Sheet ForbesTransposedSheet { get; set; }
        private Table CompanyTable { get; set; }
        private Table CompanyTransposedTable { get; set; }
        private Form  CompanyForm { get; set; }
        private Form  CompanyTransposedForm { get; set; }
        private Table CountryTable { get; set; }
        private Table IndustryTable { get; set; }
        private bool IsLoaded { get; set; }
        private bool IsUpdating { get; set; }
        private bool IsLoadedTransposed { get; set; }
        private bool IsUpdatingTransposed { get; set; }
        private CommandTests Tests { get; set; }

        public Forbes(View view)
        {
            Tests = new CommandTests((Sheet)view.Find(Binding.ViewType.Sheet, "Tests"));

            view.HookBindingFailed(_view_BindFailed, true);

            ForbesSheet = (Sheet) view.Find(Binding.ViewType.Sheet, "Forbes");
            ForbesSheet.HookClicked(LoadAllClicked, "LoadForbes", true);
            ForbesSheet.HookClicked(ClearAllClicked, "ClearForbes", true);
            ForbesSheet.HookClicked(StartUpdateClicked, "StartUpdate", true);
            ForbesSheet.HookClicked(ShowColumnClicked, "ShowColumn", true);
            ForbesSheet.HookClicked(ShowDialogClicked, "ShowDialog", true);

            CompanyTable = (Table)ForbesSheet.Find(Binding.ViewType.Table, "Company");
            CompanyTable.SelectionChanged += _companyTable_SelectionChanged;
            CompanyTable.ObjectChanged += _companyTable_ObjectChanged;
            CompanyTable.Model = new CompanyList();

            CompanyForm = (Form)ForbesSheet.Find(Binding.ViewType.Form, "Company");
            CompanyForm.ObjectChanged += _companyForm_ObjectChanged;

            ForbesTransposedSheet = (Sheet)view.Find(Binding.ViewType.Sheet, "Forbes_transposed");
            ForbesTransposedSheet.HookClicked(LoadAllClickedTransposed, "TransposedLoadForbes", true);
            ForbesTransposedSheet.HookClicked(ClearAllClickedTransposed, "TransposedClearForbes", true);
            ForbesTransposedSheet.HookClicked(StartUpdateClickedTransposed, "TransposedStartUpdate", true);
            ForbesTransposedSheet.HookClicked(ShowRowClicked, "TransposedShowRow", true);
            ForbesTransposedSheet.HookClicked(ShowDialogClickedTransposed, "TransposedShowDialog", true);

            CompanyTransposedTable = (Table)ForbesTransposedSheet.Find(Binding.ViewType.Table, "CompanyTransposed");
            CompanyTransposedTable.SelectionChanged += _companyTransposedTable_SelectionChanged;
            CompanyTransposedTable.ObjectChanged += _companyTransposedTable_ObjectChanged;
            CompanyTransposedTable.Model = new CompanyList();

            CompanyTransposedForm = (Form)ForbesTransposedSheet.Find(Binding.ViewType.Form, "CompanyTransposed");
            CompanyTransposedForm.ObjectChanged += _companyTransposedForm_ObjectChanged;

            CountryTable = (Table) view.Find(Binding.ViewType.Table, "Country");
            IndustryTable = (Table) view.Find(Binding.ViewType.Table, "Industry");
            EnableControls();
        }

        void _companyForm_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            // this is just for demo purpose, just to get the table to update, careful with 
            // recursive update
            (args.Items.First() as Company).RaiseChanged();
        }

        void _companyTransposedForm_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            // this is just for demo purpose, just to get the table to update, careful with 
            // recursive update
            (args.Items.First() as Company).RaiseChanged();
        }

        void _companyTable_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            var model = args.Items.Last();
            if (model == CompanyForm.Model)
                ((Company) model).RaiseChanged();
        }

        void _companyTable_SelectionChanged(object sender, SelectionChangedArgs args)
        {
            CompanyForm.Model = args.Items.Last();
        }

        void _companyTransposedTable_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            var model = args.Items.Last();
            if (model == CompanyTransposedForm.Model)
                ((Company)model).RaiseChanged();
        }

        void _companyTransposedTable_SelectionChanged(object sender, SelectionChangedArgs args)
        {
            CompanyTransposedForm.Model = args.Items.Last();
        }

        private void _view_BindFailed(object sender, BindingFailedEventArgs args)
        {
            MessageBox.Show(args.Exception.Message, args.View.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        void LoadAllClicked(object sender, CommandEventArgs args)
        {
            var companyList =(CompanyList) CompanyTable.Model;
            companyList.Load();
            RebindReferenceLists(companyList);
            companyList.RaiseChanged();
            IsLoaded = true;
            EnableControls();
        }

        void LoadAllClickedTransposed(object sender, CommandEventArgs args)
        {
            var companyList = (CompanyList)CompanyTransposedTable.Model;
            companyList.Load();
            RebindReferenceLists(companyList);
            companyList.RaiseChanged();
            IsLoadedTransposed = true;
            EnableControls();
        }

        void ClearAllClicked(object sender, CommandEventArgs args)
        {
            var companyList = (CompanyList)CompanyTable.Model;
            companyList.Unload();
            RebindReferenceLists(companyList);
            IsLoaded = false;
            EnableControls();
        }

        void ClearAllClickedTransposed(object sender, CommandEventArgs args)
        {
            var companyList = (CompanyList)CompanyTransposedTable.Model;
            companyList.Unload();
            RebindReferenceLists(companyList);
            IsLoadedTransposed = false;
            EnableControls();
        }

        private void RebindReferenceLists(CompanyList clist)
        {
            CountryTable.Model = clist.CountryList;
            IndustryTable.Model = clist.IndustryList;
        }

        void StartUpdateClicked(object sender, CommandEventArgs args)
        {
            var cmd = (Command) sender;
            var update = !(bool) (cmd.Value ?? false);
            cmd.Value = update;
            cmd.Caption = update ? "Stop Update" : "Start Update";
            var companyList = (CompanyList)CompanyTable.Model;
            companyList.Update(update);
            IsUpdating = update;
            EnableControls();
        }

        void StartUpdateClickedTransposed(object sender, CommandEventArgs args)
        {
            var cmd = (Command)sender;
            var update = !(bool)(cmd.Value ?? false);
            cmd.Value = update;
            cmd.Caption = update ? "Stop Update" : "Start Update";
            var companyList = (CompanyList)CompanyTransposedTable.Model;
            companyList.Update(update);
            IsUpdatingTransposed = update;
            EnableControls();
        }

        private void ShowColumnClicked(object sender, CommandEventArgs args)
        {
            var visible =CompanyTable.ToggleCategoryVisibility("Industry");
            var cmd = (Command)sender;
            cmd.Caption = visible ? "Hide Industry" : "Show Industry";
        }

        private void ShowRowClicked(object sender, CommandEventArgs args)
        {
            var visible = CompanyTransposedTable.ToggleCategoryVisibility("Industry");
            var cmd = (Command)sender;
            cmd.Caption = visible ? "Hide Industry" : "Show Industry";
        }

        private void ShowDialogClicked(object sender, CommandEventArgs args)
        {
            var v = new Forbes2000 { Model = (IEnumerable)CompanyTable.Model };
            var interop = new WindowInteropHelper(v) {Owner = App.Instance.Root.Handle};
            v.ShowDialog(); // or v.Show();
        }

        private void ShowDialogClickedTransposed(object sender, CommandEventArgs args)
        {
            var v = new Forbes2000 { Model = (IEnumerable)CompanyTransposedTable.Model };
            var interop = new WindowInteropHelper(v) { Owner = App.Instance.Root.Handle };
            v.ShowDialog(); // or v.Show();
        }

        private void EnableControls()
        {
            ForbesSheet.FindCommand("LoadForbes").IsEnabled = !IsLoaded && !IsUpdating;
            ForbesSheet.FindCommand("ClearForbes").IsEnabled = IsLoaded && !IsUpdating; 
            ForbesSheet.FindCommand("StartUpdate").IsEnabled = IsLoaded;
            ForbesTransposedSheet.FindCommand("TransposedLoadForbes").IsEnabled = !IsLoadedTransposed && !IsUpdatingTransposed;
            ForbesTransposedSheet.FindCommand("TransposedClearForbes").IsEnabled = IsLoadedTransposed && !IsUpdatingTransposed;
            ForbesTransposedSheet.FindCommand("TransposedStartUpdate").IsEnabled = IsLoadedTransposed;
        }
    }
}
