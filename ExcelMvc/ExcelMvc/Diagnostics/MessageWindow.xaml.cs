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
    using System.ComponentModel;
    using System.Windows;
    using System.Windows.Input;
    using Runtime;

    /// <summary>
    /// Implements a visual sink for exception and information messages
    /// </summary>
    public partial class MessageWindow : Window
    {
        public MessageWindow()
        {
            InitializeComponent();
            Closed += MessageWindow_Closed;
            Closing += MessageWindow_Closing;
            LayoutRoot.DataContext = Messages.Instance;
        }

        private static MessageWindow Instance { get; set; }

        /// <summary>
        /// Creates and shows to the status window
        /// </summary>
        public static void ShowInstance()
        {
            AsyncActions.Post(
                state =>
                {
                    CreateInstance();
                    // var interop = new WindowInteropHelper(Instance) { Owner = App.Instance.MainWindow.Handle };
                    Instance.Show();
                },
                null,
                false);
        }

        /// <summary>
        /// Hides the singleton
        /// </summary>
        public static void HideInstance()
        {
            AsyncActions.Post(
                state =>
                { 
                    Instance?.Hide(); 
                },
                null,
                false);
        }

        public static void CreateInstance()
        {
            Instance = Instance ?? new MessageWindow();
        }

        private static void MessageWindow_Closing(object sender, CancelEventArgs e)
        {
            if (ReferenceEquals(sender, Instance))
            {
                e.Cancel = true;
                HideInstance();
            }
        }

        private static void MessageWindow_Closed(object sender, EventArgs e)
        {
            if (ReferenceEquals(sender, Instance))
                Instance = null;
        }

        private void ButtonClear_OnClick(object sender, RoutedEventArgs e)
        {
            Messages.Instance.Clear();
        }

        private void ButtonHide_OnClick(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void LineLimit_OnKeyDown(object sender, KeyEventArgs e)
        {
            var key = Convert.ToInt32(e.Key);
            e.Handled = (key < Convert.ToInt32(Key.D0) || key > Convert.ToInt32(Key.D9))
                     && (key < Convert.ToInt32(Key.NumPad0) || key > Convert.ToInt32(Key.NumPad9));
        }
    }
}
