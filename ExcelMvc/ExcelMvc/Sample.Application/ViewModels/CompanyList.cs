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
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Threading;
using ExcelMvc.Runtime;

namespace Sample.Application.ViewModels
{
    public class CompanyList : List<Company>, INotifyCollectionChanged
    {
        public event NotifyCollectionChangedEventHandler CollectionChanged = delegate { };

        public List<string> CountryList { get; private set; }
        public List<string> IndustryList { get; private set; }

        public CompanyList()
        {
            CountryList = new List<string>();
            IndustryList = new List<string>();
        }

        public void Load()
        {
            Clear();

            var lists = new Models.CompanyList();
            lists.Load();
            foreach (var item in lists)
                Add(new Company {Model = item});

            CountryList.Clear();
            CountryList.AddRange(lists.Select(x => x.Country).Distinct());
            CountryList.Sort();

            IndustryList.Clear();
            IndustryList.AddRange(lists.Select(x => x.Industry).Distinct());
            IndustryList.Sort();
        }

        public void Unload()
        {
            Clear();
            RaiseChanged();
            CountryList.Clear();
            IndustryList.Clear();
        }

        public void RaiseChanged()
        {
            CollectionChanged(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }

        private Thread _updateThread;
        private ManualResetEvent _stopEvent;
        public void Update(bool start)
        {
            if (start)
            {
                _updateThread = new Thread(RunUpdate) {Name = RangeUpdator.NameOfAsynUpdateThread};
                _stopEvent = new ManualResetEvent(false);
                _updateThread.Start();
            }
            else if (_stopEvent != null)
            {
                _stopEvent.Set();
            }
        }

        private void RunUpdate()
        {
            var random = new Random();
            while (!_stopEvent.WaitOne(100))
            {
                var idx = (int) (random.NextDouble() * 25);
                if (idx == Count) idx--;
                var x = this[idx];
                x.Model.Profits = (0.5 - random.NextDouble())* 100;
                x.RaiseChanged("Profits");
            }
        }
    }
}
