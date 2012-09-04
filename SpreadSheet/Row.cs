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

        private Sheet sheet;

        private SortedDictionary<int, Cell> m_cells = new SortedDictionary<int, Cell>();

        /// <summary>
        /// Initializes a new instance of the <see cref="Row"/> class.
        /// </summary>
        /// <param name="sheet">Owner sheet.</param>
        /// <param name="row">Row index.</param>
		internal Row(Sheet sheet, int row)
		{
            this.sheet = sheet;
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
        		int result = int.MaxValue;
        		foreach (int idx in m_cells.Keys)
        		{
        			if (idx < result)
        			{
        				result = idx;
        			}
        		}
                foreach (CellRange cr in this.sheet.mergedCells)
                {
                    if (cr.FirstRow <= this.RowIndex && cr.LastRow >= this.RowIndex && cr.FirstColumn < result)
                    {
                        result = cr.FirstColumn;
                    }
                }
                return (result == int.MaxValue ? -1 : result);
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
                foreach (CellRange cr in this.sheet.mergedCells)
                {
                    if (cr.FirstRow <= this.RowIndex && cr.LastRow >= this.RowIndex && cr.LastColumn > result)
                    {
                        result = (uint)cr.LastColumn;
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
                    Cell nc = new Cell(this.sheet, this.row, column);
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

        /// <summary>
        /// Adjust row height to fit all cell content
        /// </summary>
        /// <param name="allRows">If false than do not override row height if already set</param>
        public void AutoSizeRowHeight(bool allRows = false)
        {
            int f = FirstCell;
            if (f == -1 || (!allRows && this.Height.HasValue))
                return;

            ushort height = 17;

            for (int i = f; i <= LastCell; ++i)
            {
                int cheight = 0;
                CellRange cr = this.sheet.mergedCells.GetAtPosition(RowIndex, i);
                if (cr != null)
                {
                    if (!(cr.LastRow == RowIndex && cr.FirstColumn == i))
                        continue;
                    
                    cheight = 0;
                    for (int r = cr.FirstRow; r < cr.LastRow; ++r)
                        cheight += (this.sheet[cr.LastRow].Height ?? 17);
                }

                ushort nheight = (ushort)Math.Max((int)this[i].CalculateCellHeight() + 2 - cheight, 0);

                height = Math.Max(height, nheight);
            }

            this.Height = height;
        }
	}
}
