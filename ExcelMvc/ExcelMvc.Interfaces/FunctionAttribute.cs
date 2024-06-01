/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia

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

namespace Function.Interfaces
{
    /// <summary>
    /// Decorates functions that are to be exported as User Defined Functions.
    /// https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public class FunctionAttribute : Attribute
    {
        /// <summary>
        /// Specifies which category that the function should be listed in the function wizard.
        /// </summary>
        public string Category;
        /// <summary>
        /// The function name as it will appear in the Function Wizard.
        /// </summary>
        public string Name;
        /// <summary>
        /// The Description of the function when it is selected in the Function Wizard.
        /// </summary>
        public string Description;
        /// <summary>
        /// The help information displayed when the Help button is clicked.
        /// It can be in either "chm-file!HelpContextID" or "https://address/path_to_file_in_site!0". 
        /// </summary>
        public string HelpTopic;

        /// <summary>
        /// Indicates the type of function, 0, 1 or 2.
        /// </summary>
        public byte FunctionType = 1;

        /// <summary>
        /// Registers the function as volatile, i.e. recalculates every time the worksheet recalculates.
        /// (pxTypeText +='!')
        /// </summary>
        public bool IsVolatile;

        /// <summary>
        /// Registers the function as macro sheet equivalent, handling uncalculated cells.
        /// pxTypeText +='#'
        /// </summary>
        public bool IsMacro;

        /// <summary>
        /// Indicates if the function is listed in the Function Wizard.
        /// </summary>
        public bool IsHidden = false;

        /// <summary>
        /// Registers the function as an asynchronous function.
        /// (pxTypeText=>(pxArgsTypeText)X)
        /// </summary>
        public bool IsAsync;

        /// <summary>
        /// Indicates the function is thread-safe.
        /// (pxTypeTex +='$')
        /// </summary>
        public bool IsThreadSafe;

        /// <summary>
        /// Indicates the function is cluster-safe.
        /// (pxTypeText += '&')
        /// </summary>
        public bool IsClusterSafe;

        public FunctionAttribute() { }
        public FunctionAttribute(string description)
            => Description = description;
    }
}
