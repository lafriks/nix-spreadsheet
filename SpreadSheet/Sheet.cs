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
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Drawing;

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Sheet.
	/// </summary>
	public class Sheet : IEnumerable<Row>
	{
		private SpreadSheetDocument document;

		/// <summary>
		/// Maximal row count.
		/// </summary>
		public const int MaxRows = 65536;
		/// <summary>
		/// Maximal column count.
		/// </summary>
        public const int MaxColumns = 256;

        private SortedDictionary<int, Row> m_rows = new SortedDictionary<int, Row>();
        
        private ColumnList columns;

        /// <summary>
        /// Sheet columns.
        /// </summary>
        public ColumnList Columns
        { 
        	get { return columns; }
        }
        
		/// <summary>
		/// Gets the owner document.
		/// </summary>
		/// <value>The owner document.</value>
		public SpreadSheetDocument Document
		{
			get { return document; }
		}

		/// <summary>
		/// Last used row index.
		/// </summary>
        public uint LastRow
        {
            
            get
            {
                uint result = 0;
                foreach (int key in m_rows.Keys)
                {
                    if (key > result)
                    {
                        result = (uint)key;
                    }
                }
                return result;
            }
        }

        /// <summary>
        /// First used row index.
        /// </summary>
        public uint FirstRow
        {
            
            get
            {
            	uint result = m_rows.Count > 0 ? uint.MaxValue : 0;
                foreach (int key in m_rows.Keys)
                {
                    if (key < result)
                    {
                        result = (uint)key;
                    }
                }
                return result;
            }
        }

        /// <summary>
        /// First used column index.
        /// </summary>
        public uint FirstColumn
        {
        	get
        	{
        		int result = -1;
        		
        		foreach (Row row in m_rows.Values)
        		{
        			if (row.FirstCell != -1 && (result == -1 || result > row.FirstCell))
        			{
        				result = row.FirstCell;
        			}
        		}
        		
        		foreach (Column col in columns)
        		{
        			if (result == -1 || result > col.ColumnIndex)
        			{
        				result = col.ColumnIndex;
        			}
        		}
        		
        		return result < 0 ? 0 : (uint)result;
        	}
        }

        /// <summary>
        /// Last used column.
        /// </summary>
        public uint LastColumn
        {
        	get
        	{
        		uint result = 0;
        		
        		foreach (Row row in m_rows.Values)
        		{
        			if (result < row.LastCell)
        			{
        				result = row.LastCell;
        			}
        		}
        		
        		foreach (Column col in columns)
        		{
        			if (result < col.ColumnIndex)
        			{
        				result = (uint)col.ColumnIndex;
        			}
        		}
        		
        		return result;
        	}
        }

		/// <summary>
		/// Initializes a new instance of the <see cref="Sheet"/> class.
		/// </summary>
		/// <param name="document">The owner document.</param>
		/// <param name="name">The name.</param>
		public Sheet(SpreadSheetDocument document, string name)
		{
			this.document = document;
			this.name = name;
			this.columns = new ColumnList(this);
		}

		private string name = "Sheet";

		/// <summary>
		/// Gets or sets the sheet name.
		/// </summary>
		/// <value>The sheet name.</value>
		public string Name
		{
			get { return name; }
			set
			{
				this.Document.ChangeSheetName(name, value);
				name = value;
			}
		}

		/// <summary>
		/// Get row.
		/// </summary>
		public Row this[int row]
		{
			get
			{
                if (row >= MaxRows || row < 0)
                	throw new ArgumentOutOfRangeException("row");
                if ( this.m_rows.ContainsKey(row) )
                	return this.m_rows[row];
                else
                {
                	Row nr = new Row(row);
                	this.m_rows.Add(row, nr);
                	return nr;
                }
                
			}
		}

		/// <summary>
		/// Get cell by row and column index.
		/// </summary>
        public Cell this[int row, int column]
        {
            get
            {
            	return this[row][column];
            }
        }

        /// <summary>
        /// Get cell by name.
        /// </summary>
        public Cell this[string name]
        {
            get
            {
            	int r;
            	int c;
            	Utils.ParseCellName(name, out r, out c);
            	return this[r, c];
            }
        }
        /// <summary>
        /// Get cell range.
        /// </summary>
        /// <param name="range">Range in format (ex. A1:B3)</param>
        /// <returns>Cell range.</returns>
        public CellRange GetCellRange(string range)
        {
        	int fr, fc, lr, lc, t;
        	string[] rng = range.Split(':');
        	if ( rng.GetLength(0) == 1 )
        	{
        		Utils.ParseCellName(rng[0], out fr, out fc);
        		lr = fr;
        		lc = fc;
        	}
        	else
        	{
        		Utils.ParseCellName(rng[0], out fr, out fc);
        		Utils.ParseCellName(rng[1], out lr, out lc);
        	}
        	if ( fr > lr )
        	{
        		t = lr;
        		lr = fr;
        		fr = t;
        	}
        	if ( fc > lc )
        	{
        		t = lc;
        		lc = fc;
        		fc = t;
        	}
        	return GetCellRange(fr, fc, lr, lc);
        }

        /// <summary>
        /// Get cell range.
        /// </summary>
        /// <param name="firstRow">First row.</param>
        /// <param name="firstColumn">First column.</param>
        /// <param name="lastRow">Last row.</param>
        /// <param name="lastColumn">Last column.</param>
        /// <returns></returns>
        public CellRange GetCellRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
        	return new CellRange(this, firstRow, firstColumn, lastRow, lastColumn);
        }

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.m_rows.Values.GetEnumerator();
		}

		IEnumerator<Row> IEnumerable<Row>.GetEnumerator()
		{
			return this.m_rows.Values.GetEnumerator();
		}

		#region DataTable helper
		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstCell">Cell name.</param>
		/// <param name="dataView">Data view to get data from.</param>
		public void InsertTable(string firstCell, DataView dataView)
		{
			InsertTable(firstCell, dataView, null);
		}

		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstCell">Cell name.</param>
		/// <param name="dataView">Data view to get data from.</param>
		/// <param name="columnHeaders">Column header dictionary, where key is columnName and value is header text.</param>
		public void InsertTable(string firstCell, DataView dataView, Dictionary<string, string> columnHeaders)
		{
			InsertTable(firstCell, dataView, columnHeaders, true);
		}

		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstCell">Cell name.</param>
		/// <param name="dataView">Data view to get data from.</param>
		/// <param name="columnHeaders">Column header dictionary, where key is columnName and value is header text.</param>
		/// <param name="showHeader">Show table header.</param>
		public void InsertTable(string firstCell, DataView dataView, Dictionary<string, string> columnHeaders, bool showHeader)
		{
			InsertTable(firstCell, dataView, columnHeaders, showHeader, true);
		}

		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstCell">Cell name.</param>
		/// <param name="dataView">Data view to get data from.</param>
		/// <param name="columnHeaders">Column header dictionary, where key is columnName and value is header text.</param>
		/// <param name="showHeader">Show table header.</param>
		/// <param name="formatTable">Format table with borders.</param>
		public void InsertTable(string firstCell, DataView dataView, Dictionary<string, string> columnHeaders, bool showHeader, bool formatTable)
		{
			int fc, fr;
			Utils.ParseCellName(firstCell, out fr, out fc);
			InsertTable(fr, fc, dataView, columnHeaders, showHeader, formatTable);
		}

		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstRow">First row index.</param>
		/// <param name="firstColumn">First column index.</param>
		/// <param name="dataView">Data view to get data from.</param>
		public void InsertTable(int firstRow, int firstColumn, DataView dataView)
		{
			InsertTable(firstRow, firstColumn, dataView, null);
		}

		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstRow">First row index.</param>
		/// <param name="firstColumn">First column index.</param>
		/// <param name="dataView">Data view to get data from.</param>
		/// <param name="columnHeaders">Column header dictionary, where key is columnName and value is header text.</param>
		public void InsertTable(int firstRow, int firstColumn, DataView dataView, Dictionary<string, string> columnHeaders)
		{
			InsertTable(firstRow, firstColumn, dataView, columnHeaders, true);
		}

		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstRow">First row index.</param>
		/// <param name="firstColumn">First column index.</param>
		/// <param name="dataView">Data view to get data from.</param>
		/// <param name="columnHeaders">Column header dictionary, where key is columnName and value is header text.</param>
		/// <param name="showHeader">Show table header.</param>
		public void InsertTable(int firstRow, int firstColumn, DataView dataView, Dictionary<string, string> columnHeaders, bool showHeader)
		{
			InsertTable(firstRow, firstColumn, dataView, columnHeaders, showHeader, true);
		}

		/// <summary>
		/// Insert DataTable view into sheet starting at specified cell.
		/// </summary>
		/// <param name="firstRow">First row index.</param>
		/// <param name="firstColumn">First column index.</param>
		/// <param name="dataView">Data view to get data from.</param>
		/// <param name="columnHeaders">Column header dictionary, where key is columnName and value is header text.</param>
		/// <param name="showHeader">Show table header.</param>
		/// <param name="formatTable">Format table with borders.</param>
		public void InsertTable(int firstRow, int firstColumn, DataView dataView, Dictionary<string, string> columnHeaders, bool showHeader, bool formatTable)
		{
			if (columnHeaders == null)
			{
				columnHeaders = new Dictionary<string, string>();
				foreach (DataColumn col in dataView.Table.Columns)
				{
					columnHeaders.Add(col.ColumnName, col.ColumnName);
				}
			}

			int c = 0;
			if (showHeader)
			{
				foreach (string columnName in columnHeaders.Keys)
				{
					this[firstRow, c].Value = columnHeaders[columnName];
					c++;
				}
			}

			for (int r = 0; r < dataView.Count; r++)
			{
				c = 0;
				foreach (string columnName in columnHeaders.Keys)
				{
					this[r + firstRow + (showHeader ? 1 : 0), c + firstColumn].Value = dataView[r][columnName];
					c++;
				}
			}

			if (formatTable)
			{
				this.GetCellRange(firstRow, firstColumn, firstRow + dataView.Count - (showHeader ? 0 : 1), firstColumn + columnHeaders.Count - 1)
							.DrawTable(Color.Black, BorderLineStyle.Thin, BorderLineStyle.Medium);
				if (showHeader)
				{
					this.GetCellRange(firstRow, firstColumn, firstRow, firstColumn + columnHeaders.Count - 1)
							.DrawBorder(Color.Black, BorderLineStyle.Medium)
							.SetAlignment(CellHorizontalAlignment.Centred, CellVerticalAlignment.Centred)
							.SetBackground(Color.Gray);
				}
			}
		}
		#endregion
	}
}
