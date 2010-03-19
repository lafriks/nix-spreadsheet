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

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Sheet.
	/// </summary>
	public class ColumnList : IEnumerable<Column>
	{
	
		private SortedDictionary<int, Column> m_columns = new SortedDictionary<int, Column>();
		
		public Column this[int col]
		{
			get
			{
                if (col >= Sheet.MaxColumns || col < 0)
                	throw new ArgumentOutOfRangeException("column");
                if ( this.m_columns.ContainsKey(col) )
                	return this.m_columns[col];
                else
                {
                	Column nr = new Column(col);
                	this.m_columns.Add(col, nr);
                	return nr;
                }
                
			}
		}
		
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.m_columns.Values.GetEnumerator();
		}
		
		IEnumerator<Column> IEnumerable<Column>.GetEnumerator()
		{
			return this.m_columns.Values.GetEnumerator();
		}
		
	}
}