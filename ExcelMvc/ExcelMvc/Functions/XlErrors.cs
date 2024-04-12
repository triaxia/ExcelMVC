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

using System.Collections.Generic;
using System.Linq;

namespace ExcelMvc.Functions
{
    public enum XlErrors
    {
        xlerrUnknown = -1,
        xlerrNull = 0,
        xlerrDiv0 = 7,
        xlerrValue = 15,
        xlerrRef = 23,
        xlerrName = 29,
        xlerrNum = 36,
        xlerrNA = 42,
        xlerrGettingData = 43
    }

    public interface XlError
    {
        XlErrors Type { get; }
    }

    public class XlErrorKnown : XlError
    {
        private XlErrorKnown() { }
        public XlErrors Type => XlErrors.xlerrUnknown;
        public static readonly XlErrorKnown Instance = new XlErrorKnown();
        public override string ToString() => "#???!";
    }

    public class XlErrorNull : XlError
    {
        private XlErrorNull() { }
        public XlErrors Type => XlErrors.xlerrNull;
        public static readonly XlErrorNull Instance = new XlErrorNull();
        public override string ToString() => "#NULL!";
    }

    public class XlErrorDiv0 : XlError
    {
        private XlErrorDiv0() { }
        public XlErrors Type => XlErrors.xlerrDiv0;
        public static readonly XlErrorDiv0 Instance = new XlErrorDiv0();
        public override string ToString() => "#DIV0!";
    }

    public class XlErrorValue : XlError
    {
        private XlErrorValue() { }
        public XlErrors Type => XlErrors.xlerrValue;
        public static readonly XlErrorValue Instance = new XlErrorValue();
        public override string ToString() => "#VALUE!";
    }

    public class XlErrorRef : XlError
    {
        private XlErrorRef() { }
        public XlErrors Type => XlErrors.xlerrRef;
        public static readonly XlErrorRef Instance = new XlErrorRef();
        public override string ToString() => "#REF!";
    }

    public class XlErrorName : XlError
    {
        private XlErrorName() { }
        public XlErrors Type => XlErrors.xlerrName;
        public static readonly XlErrorName Instance = new XlErrorName();
        public override string ToString() => "#NAME?";
    }

    public class XlErrorNum: XlError
    {
        private XlErrorNum() { }
        public XlErrors Type => XlErrors.xlerrNum;
        public static readonly XlErrorNum Instance = new XlErrorNum();
        public override string ToString() => "#NUM!";
    }

    public class XlErrorNA : XlError
    {
        private XlErrorNA() { }
        public XlErrors Type => XlErrors.xlerrNA;
        public static readonly XlErrorNA Instance = new XlErrorNA();
        public override string ToString() => "#N/A";
    }

    public class XlErrorGettingData : XlError
    {
        private XlErrorGettingData() { }
        public XlErrors Type => XlErrors.xlerrGettingData;
        public static readonly XlErrorGettingData Instance = new XlErrorGettingData();
        public override string ToString() => "#Data!";
    }

    public static class XlErrorFactory
    {
        private static Dictionary<XlErrors, XlError> Errors;
        static XlErrorFactory()
        {
            Errors = new Dictionary<XlErrors, XlError>()
             {
                 {XlErrors.xlerrNull, XlErrorNull.Instance},
                 {XlErrors.xlerrDiv0, XlErrorDiv0.Instance},
                 {XlErrors.xlerrValue, XlErrorValue.Instance},
                 {XlErrors.xlerrRef, XlErrorRef.Instance},
                 {XlErrors.xlerrName, XlErrorName.Instance},
                 {XlErrors.xlerrNum, XlErrorNum.Instance},
                 {XlErrors.xlerrNA, XlErrorNA.Instance},
                 {XlErrors.xlerrGettingData, XlErrorGettingData.Instance}
             };
        }
        public static XlError TypeToObject(XlErrors type)
            => Errors.TryGetValue(type, out var value) ? value : Errors[XlErrors.xlerrUnknown];
        public static XlErrors ObjectToType(XlError obj)
            => Errors.Single(x => x.Value == obj).Key;
    }
}
