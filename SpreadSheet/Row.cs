/*
 * Library for generating spreadsheet documents.
 * Copyright (C) 2008, Lauris Bukðis-Haberkorns <lauris@nix.lv>
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
using System.Collections;
using System.Collections.Generic;

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Description of Row.
	/// </summary>
	public class Row : IEnumerable<Cell>
	{
		private int row;

        private SortedDictionary<int, Cell> m_cells = new SortedDictionary<int, Cell>();

        /// <summary>
        /// Initializes a new instance of the <see cref="Row"/> class.
        /// </summary>
        /// <param name="row">Row index.</param>
		public Row(int row)
		{
			this.row = row;
		}

		/// <summary>
		/// Row index.
		/// </summary>
        public int RowIndex
        {
            get
            {
                return this.row;
            }
        }

        internal bool HasCellInColumn(int column)
        {
            return m_cells.ContainsKey(column);
        }

        internal bool RemoveCellAtColumn(int column)
        {
            return m_cells.Remove(column);
        }
        
        private ushort? height = null;

        /// <summary>
        /// Row height in pixels
        /// </summary>
        public ushort? Height
        {
            get { return height; }
            set { height = value; }
        }

        /// <summary>
        /// First used cell in row.
        /// </summary>
        public int FirstCell
        {
        	get
        	{
        		int result = (m_cells.Count == 0 ? -1 : int.MaxValue);
        		foreach (int idx in m_cells.Keys)
        		{
        			if (idx < result)
        			{
        				result = idx;
        			}
        		}
        		return result;
        	}
        }

        /// <summary>
        /// Last used cell in row
        /// </summary>
        public uint LastCell
        {
        	get
        	{
        		uint result = 0;
        		foreach (int idx in m_cells.Keys)
        		{
        			if (idx > result)
        			{
        				result = (uint)idx;
        			}
        		}
        		return result;
        	}
        }

        /// <summary>
        /// Get cell by column index in current row.
        /// </summary>
		public Cell this[int column]
		{
			get
			{
                if (column >= Sheet.MaxColumns || column < 0)
                    throw new ArgumentOutOfRangeException("column");
                if ( this.m_cells.ContainsKey(column) )
                	return this.m_cells[column];
                else
                {
                    Cell nc = new Cell(this.row, column);
                    this.m_cells.Add(column, nc);
                    return nc;
                }
			}
		}

		public IEnumerator<Cell> GetEnumerator()
		{
			return this.m_cells.Values.GetEnumerator();
		}
		
		IEnumerator IEnumerable.GetEnumerator()
		{
			return this.m_cells.Values.GetEnumerator();
		}
	}
}
