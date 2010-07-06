/*
 * Library for generating spreadsheet documents.
 * Copyright (C) 2010, Lauris Bukšis-Haberkorns <lauris@nix.lv>
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace Nix.SpreadSheet
{
    public class ColumnRange
    {
        private int startColumn;

        public int StartColumn
        {
            get { return startColumn; }
        }

        private int endColumn;

        public int EndColumn
        {
            get { return endColumn; }
        }

        private Sheet sheet;

        public Sheet Sheet
        {
            get { return sheet; }
        }

        public ColumnRange(Sheet sheet, int startColumn, int endColumn)
        {
            this.sheet = sheet;
            this.startColumn = startColumn;
            this.endColumn = endColumn;
        }

        public uint Width
        {
            set
            {
                for (int i = this.StartColumn; i <= this.EndColumn; i++)
                {
                    this.sheet.Columns[i].Width = value;
                }
            }
            get
            {
                uint width = this.Sheet.Columns[this.StartColumn].Width;
                for (int i = this.StartColumn + 1; i <= this.EndColumn; i++)
                {
                    if (this.sheet.Columns[i].Width != width)
                        throw new Exception("ColumnRange columns does not have same width");
                }
                return width;
            }
        }
    }
}
