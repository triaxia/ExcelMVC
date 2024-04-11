/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany (2013)

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

namespace ExcelMvc.Diagnostics
{
    using System;
    using System.Collections.Concurrent;
    using System.ComponentModel;
    using System.Linq;

    public class Messages : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public string Error =>
            string.Join(System.Environment.NewLine, ErrorLines.ToArray().Reverse());

        public string Info =>
            string.Join(System.Environment.NewLine, InfoLines.ToArray().Reverse());

        public int LineLimit { get; set; }

        private ConcurrentQueue<string> ErrorLines { get; }
            = new ConcurrentQueue<string>();
        public ConcurrentQueue<string> InfoLines { get; }
            = new ConcurrentQueue<string>();

        public static readonly Messages Instance = new Messages();

        public Messages()
        {
            LineLimit = 2000;
        }

        public void Clear()
        {
            while (ErrorLines.TryDequeue(out var _)) ;
            while (InfoLines.TryDequeue(out var _)) ;
            RaiseErrorChanged();
            RaiseInfoChanged();
        }

        public void AddErrorLine(Exception ex)
        {
            AddErrorLine($"{ex}");
        }

        public void AddErrorLine(string message)
        {
            ErrorLines.Enqueue($"{DateTime.Now:O} {message}");
            while (ErrorLines.Count > LineLimit) ErrorLines.TryDequeue(out var _);
            RaiseErrorChanged();
        }

        public void AddInfoLine(string message)
        {
            InfoLines.Enqueue($"{DateTime.Now:O} {message}");
            while (InfoLines.Count > LineLimit) InfoLines.TryDequeue(out var _);
            RaiseInfoChanged();
        }

        private void RaiseErrorChanged()
        {
            PropertyChanged(this, new PropertyChangedEventArgs("Error"));
        }

        private void RaiseInfoChanged()
        {
            PropertyChanged(this, new PropertyChangedEventArgs("Info"));
        }
    }
}
