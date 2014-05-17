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
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Sample.Models
{
    public class CompanyList : List<Company>
    {
        private struct HeadingIndices
        {
            public int IndexOfCompany;
            public int IndexOfIndustry;
            public int IndexOfCountry;
            public int IndexOfMarket;
            public int IndexOfSales;
            public int IndexOfProfits;
            public int IndexOfAssets;
            public int IndexOfRank;
        }

        public void Load()
        {
            Clear();
            var path = Assembly.GetExecutingAssembly().Location;
            path = path == "" ? AppDomain.CurrentDomain.BaseDirectory : Path.GetDirectoryName(path);
            path = Path.Combine(path, "Forbes.csv");
            var lines = File.ReadAllLines(path);
            var indices = ToIndices(ParseLine(lines[0]));
            for (var idx = 1; idx < lines.Length; idx++)
            {
                var cells = ParseLine(lines[idx]);
                var company = ToCompany(cells, indices);
                if (company != null)
                    Add(company);
            }
        }

        private static readonly Regex CellPattern = new Regex("([^\",]*,)|(\"(([^\"]*)|(\"\"))*\",)");
        private static List<string> ParseLine(string line)
        {
            Match match;
            var cells = new List<string>();
            var pos = 0;
            var length = line.Length;
            while (pos < length && (match = CellPattern.Match(line, pos)).Success)
            {
                cells.Add(line.Substring(match.Index, match.Length - 1));
                pos = match.Index + match.Length;
            }

            if (pos < length)
                cells.Add(line.Substring(pos, length - pos));

            for (var idx = 0; idx < cells.Count; idx++)
            {
                var cell = cells[idx].Trim();
                if (cell.StartsWith("\"") && cell.EndsWith("\""))
                    cell = cell.Substring(1, cell.Length - 2).Replace("\"\"", "\"");
                cells[idx] = cell;
            }

            return cells;
        }

        private static HeadingIndices ToIndices(List<string> cells)
        {
            HeadingIndices indices;
            indices.IndexOfCompany = cells.IndexOf("Company");
            indices.IndexOfIndustry = cells.IndexOf("Industry");
            indices.IndexOfCountry = cells.IndexOf("Country");
            indices.IndexOfMarket = cells.IndexOf("Market Value");
            indices.IndexOfSales = cells.IndexOf("Sales");
            indices.IndexOfProfits = cells.IndexOf("Profits");
            indices.IndexOfAssets = cells.IndexOf("Assets");
            indices.IndexOfRank = cells.IndexOf("Rank");
            return indices;
        }

        private static Company ToCompany(List<string> cells, HeadingIndices indices)
        {
            var company = new Company
            {
                Name = cells[indices.IndexOfCompany],
                Industry = cells[indices.IndexOfIndustry],
                Country = cells[indices.IndexOfCountry],
                MarketValue = ToDouble(cells[indices.IndexOfMarket]),
                Profits = ToDouble(cells[indices.IndexOfProfits]),
                Sales = ToDouble(cells[indices.IndexOfSales]),
                Assets = ToDouble(cells[indices.IndexOfAssets]),
                Rank = ToInt(cells[indices.IndexOfRank])
            };
            return company;
        }

        private static double ToDouble(string value)
        {
            double result;
            double.TryParse(value, out result);
            return result;
        }

        private static int ToInt(string value)
        {
            int result;
            int.TryParse(value, out result);
            return result;
        }
    }
}
