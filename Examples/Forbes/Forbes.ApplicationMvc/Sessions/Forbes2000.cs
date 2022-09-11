/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany

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

namespace Forbes.Application.Sessions
{
    using System.Collections;
    using System.Linq;
    using System.Windows.Interop;
    using ExcelMvc.Controls;
    using ExcelMvc.Views;
    using Models;
    using ViewModels;
    using Views;

    internal class Forbes2000
    {
        public Forbes2000(View book, View sheet, Settings settings, string companyTableName, string companyFormName)
        {
            Sheet = sheet;
            Sheet.Unprotecting += Sheet_Unprotecting;
            Settings = settings;

            Sheet.FindCommand("ExcelMvc.Command.LoadForbes").Model = null;

            Sheet.HookClicked(LoadAllClicked, "LoadForbes", true);
            Sheet.HookClicked(ClearAllClicked, "ClearForbes", true);
            Sheet.HookClicked(StartUpdateClicked, "StartUpdate", true);
            Sheet.HookClicked(ShowColumnClicked, "ShowIndustry", true);
            Sheet.HookClicked(ShowDialogClicked, "ShowDialog", true);

            CompanyTable = (Table)Sheet.Find(ViewType.Table, companyTableName);
            CompanyTable.SelectionChanged += CompanyTable_SelectionChanged;
            CompanyTable.ObjectChanged += CompanyTable_ObjectChanged;
            CompanyTable.Model = new CompanyList();

            CompanyForm = (Form)Sheet.Find(ViewType.Form, companyFormName);
            CompanyForm.ObjectChanged += CompanyForm_ObjectChanged;

            CountryTable = (Table)book.Find(ViewType.Table, "ExcelMvc.Table.Country");
            IndustryTable = (Table)book.Find("Table.Industry");

            CompanyFilterTable = (Table)book.Find(ViewType.Table, "CompanyFilters");
            CompanyFilterTable.Model = new CompanyFilterList(CompanyFilterTable.MaxItemsToBind);

            EnableControls();
        }

        private void Sheet_Unprotecting(object sender, ViewEventArgs args)
        {
            // password protected update is slow...
            args.Accept(null);
        }

        private View Sheet
        {
            get; set;
        }

        private Form CompanyForm
        {
            get; set;
        }

        private Table CompanyTable
        {
            get; set;
        }

        private Table IndustryTable
        {
            get; set;
        }

        private Table CountryTable
        {
            get; set;
        }

        private bool IsLoaded
        {
            get; set;
        }

        private bool IsUpdating
        {
            get; set;
        }

        private Table CompanyFilterTable
        {
            get; set;
        }

        private Settings Settings
        {
            get;
            set;
        }

        private void ClearAllClicked(object sender, CommandEventArgs args)
        {
            var companyList = (CompanyList)CompanyTable.Model;
            companyList.Unload();
            RebindReferenceLists(companyList);
            IsLoaded = false;
            EnableControls();
        }

        private void CompanyForm_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            foreach (var item in args.Items.Cast<Company>())
                item.RaiseChanged(args.Paths);
        }

        private void CompanyTable_ObjectChanged(object sender, ObjectChangedArgs args)
        {
            var model = args.Items.Last();
            if (model == CompanyForm.Model)
                ((Company)model).RaiseChanged(args.Paths);
        }

        private void CompanyTable_SelectionChanged(object sender, SelectionChangedArgs args)
        {
            CompanyForm.Model = args.Items.Last();
        }

        private void EnableControls()
        {
            Sheet.FindCommand("LoadForbes").IsEnabled = !IsLoaded && !IsUpdating;
            Sheet.FindCommand("ClearForbes").IsEnabled = IsLoaded && !IsUpdating;
            Sheet.FindCommand("StartUpdate").IsEnabled = IsLoaded;
        }

        private void LoadAllClicked(object sender, CommandEventArgs args)
        {
            var companyList = (CompanyList)CompanyTable.Model;
            var filters = (CompanyFilterList)CompanyFilterTable.Model;
            companyList.Load(filters);
            RebindReferenceLists(companyList);
            companyList.RaiseChanged();
            IsLoaded = true;
            EnableControls();
        }

        private void RebindReferenceLists(CompanyList clist)
        {
            CountryTable.Model = clist.CountryList;
            IndustryTable.Model = clist.IndustryList;
        }

        private void ShowColumnClicked(object sender, CommandEventArgs args)
        {
            var visible = CompanyTable.ToggleCategoryVisibility("Industry");
            var cmd = (Command)sender;
            cmd.Caption = visible ? "Hide Industry" : "Show Industry";
        }

        private void ShowDialogClicked(object sender, CommandEventArgs args)
        {
            var v = new Forbes2000View { Model = (IEnumerable)CompanyTable.Model };
            _ = new WindowInteropHelper(v) { Owner = App.Instance.MainWindow.Handle };
            v.ShowDialog(); // or v.Show() for a floating dialog;
        }

        private void StartUpdateClicked(object sender, CommandEventArgs args)
        {
            var cmd = (Command)sender;
            var update = !(bool)(cmd.Value ?? false);
            cmd.Value = update;
            cmd.Caption = update ? "Stop Update" : "Start Update";
            var companyList = (CompanyList)CompanyTable.Model;
            companyList.Update(update, Settings);
            IsUpdating = update;
            EnableControls();
        }
    }
}
