/*
 * Library for generating spreadsheet documents.
 * Copyright (C) 2008, Lauris Bukšis-Haberkorns <lauris@nix.lv>
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

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Summary description for Cell.
    /// </summary>
    public class Column
    {
		private Sheet sheet;

		// Column of the value
        private int columnIndex;

        public Column (Sheet sheet, int col)
        {
        	this.sheet = sheet;
            this.columnIndex = col;
        }

        public int ColumnIndex
        {
            get { return this.columnIndex; }
            set { this.columnIndex = value; }
        }

        private CellStyle formatting = null;

        /// <summary>
        /// Gets or sets cell style.
        /// </summary>
        public CellStyle Formatting
        {
            get
            {
                if (this.formatting == null)
                    this.formatting = new CellStyle();
                return this.formatting;
            }
            set
            {
                this.formatting = value;
            }
        }

        private uint width = 2560;

        /// <summary>
        /// Column Width in pixels
        /// </summary>
        public uint Width
        {
            get { return this.width; }
            set { this.width = Convert.ToUInt32(Math.Round((double)value/7 * 256)); }
        }
    }
}
