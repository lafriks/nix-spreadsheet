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
    /// Cell.
    /// </summary>
    public class Cell : AbstractCell
    {
        // Row of the value
        private int row;
        // Column of the value
        private int column;

        /// <summary>
        /// Creates instance of <see cref="Cell">Cell</see> class.
        /// </summary>
        /// <param name="row">Row position</param>
        /// <param name="col">Column position</param>
        public Cell (int row, int col)
        {
            this.column = col;
            this.row = row;
        }

        /// <summary>
        /// Creates instance of <see cref="Cell">Cell</see> class.
        /// </summary>
        /// <param name="name">Cell name.</param>
        public Cell (string name)
        {
            Utils.ParseCellName(name, out this.row, out this.column);
        }

        /// <summary>
        /// Cell row index.
        /// </summary>
        public int RowIndex
        {
            get
            {
                return this.row;
            }
        }

        /// <summary>
        /// Cell column index.
        /// </summary>
        public int ColumnIndex
        {
            get
            {
                return this.column;
            }
        }

        /// <summary>
        /// Cell name.
        /// </summary>
        public string Name
        {
            get
            {
                return Utils.CellName(this.column, this.row);
            }
        }
    }
}
