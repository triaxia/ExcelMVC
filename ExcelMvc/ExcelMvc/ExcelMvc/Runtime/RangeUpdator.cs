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
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelMvc.Extensions;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Runtime
{
    public class RangeUpdator
    {
        private static readonly Lazy<RangeUpdator> Instane 
            = new Lazy<RangeUpdator>(() => new RangeUpdator(), true);
        public static RangeUpdator Instance { get { return Instane.Value; } }

        private RangeUpdator()
        {
        }

        public void Update(Range range, int rowOffset, int rows, int columnOffset, int columns, object value)
        {
            if (IsCurrentThreadExcel())
                range.MakeRange(rowOffset, rows, columnOffset, columns).Value = value;
            else
                Enqueue(new Item { Range = range, RowOffset = rowOffset, Rows = rows, ColumnOffset = columnOffset, Columns = columns , Value = value });
        }

        public void Update(Range range, Range rowIdStart, int rowCount, string rowId, int rows, int columnOffset, int columns, object value)
        {
            if (IsCurrentThreadExcel())
                range.MakeRange(RowOffsetFromRowId(rowIdStart, rowCount, rowId), rows, columnOffset, columns).Value = value;
            else
                Enqueue(new Item { Range = range, RowIdStart = rowIdStart, RowId = rowId, RowCount = rowCount,
                    Rows = rows, ColumnOffset = columnOffset, Columns = columns, Value = value });
        }

        public class Item
        {
            public Range Range { get; set; }
            public int RowOffset { get; set; }
            public int Rows { get; set; }
            public int ColumnOffset { get; set; }
            public int Columns { get; set; }
            public object Value { get; set; }

            public Range RowIdStart { get; set; }
            public string RowId { get; set; }
            public int RowCount { get; set; }
        }
        private readonly Queue<Item> _items = new Queue<Item>();

        [MethodImpl(MethodImplOptions.Synchronized)]
        private void Enqueue(Item item)
        {
            _items.Enqueue(item);
            Start();
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        private Item Dequeue()
        {
            return _items.Count == 0 ? null : _items.Dequeue();
        }

        private Thread Worker { get; set; }
        [MethodImpl(MethodImplOptions.Synchronized)]
        private void Start()
        {
            if (Worker == null)
            {
                Worker = new Thread(Process) {IsBackground = true};
                Worker.Start();
            }
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        private void Stop()
        {
            Worker = null;
        }

        private void Process()
        {
            Item item;
            while ((item = Dequeue()) != null)
            {
                if (!Update(item))
                {
                    Enqueue(item);
                    Thread.Sleep(100);
                }
            }
            Stop();
        }

        private static bool Update(Item item)
        {
            var status = ActionExtensions.Try(() =>
            {
                var rowOffset = item.RowIdStart == null ? item.RowOffset
                    : RowOffsetFromRowId(item.RowIdStart, item.RowCount, item.RowId);
                item.Range.MakeRange(rowOffset, item.Rows, item.ColumnOffset, item.Columns).Value = item.Value;
            });

            if (status == null)
                return true;

            var exp = status as COMException;
            if (exp != null)
            {
                var errorCode = (uint) exp.ErrorCode;
                if (errorCode == 0x8001010A || errorCode == 0x800AC472)
                    return false;
            }

            //TODO
            return false;
        }

        private static int RowOffsetFromRowId(Range start, int count, string rowId)
        {
            var column = start.MakeRange(0, count, 0, 1);
            for (var idx = 0; idx < count; idx++)
            {
                if (column.Cells[idx + 1, 1].ID == rowId)
                    return idx;
            }
            return -1;
        }

        public static bool IsCurrentThreadExcel()
        {
            var threadName = Thread.CurrentThread.Name;
            return !string.IsNullOrEmpty(threadName) && threadName.CompareOrdinalIgnoreCase("VSTA_Main") == 0;
        }
    }
}
