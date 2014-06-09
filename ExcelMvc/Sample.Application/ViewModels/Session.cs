#region Header
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
#endregion Header

namespace Sample.Application.ViewModels
{
    using System.Collections.Generic;
    using System.Windows.Forms;

    using ExcelMvc.Bindings;
    using ExcelMvc.Runtime;
    using ExcelMvc.Views;
    using View = ExcelMvc.Views.View;

    public class Session : ISession
    {
        #region Fields

        private const string ViewName = "Forbes2000";

        private static readonly Dictionary<View, object> Views = new Dictionary<View, object>();

        #endregion Fields

        #region Constructors

        static Session()
        {
            App.Instance.Opening += Book_Opening;
            App.Instance.Opened += Book_Opened;

            App.Instance.Closing += Book_Closing;
            App.Instance.Closed += Book_Closed;
        }

        #endregion Constructors

        #region Methods

        public void Dispose()
        {
        }

        private static void Book_Closed(object sender, ViewEventArgs args)
        {
            // remove the applicaton model for the book closed
            Views.Remove(args.View);
        }

        private static void Book_Closing(object sender, ViewEventArgs args)
        {
        }

        private static void Book_Opened(object sender, ViewEventArgs args)
        {
            // create the application model for the book opened
            if (args.View.Id == ViewName)
                Views[args.View] = new Forbes(args.View);
        }

        private static void Book_Opening(object sender, ViewEventArgs args)
        {
            // cancel out if the book being opened is not "Forbes2000", whose view id is
            // defined by the Custom Document Propety named "ExcelMvc".
            if (args.View.Id != ViewName)
                args.Cancel();
            else
                args.View.BindingFailed += View_BindingFailed;
        }

        private static void View_BindingFailed(object sender, BindingFailedEventArgs args)
        {
            MessageBox.Show(args.Exception.Message, args.View.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        #endregion Methods
    }
}