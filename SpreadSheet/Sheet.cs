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
	public class Sheet : IEnumerable<Row>
	{
		private SpreadSheetDocument document;

		public const int MaxRows = 65536;
        public const int MaxColumns = 256;

        private SortedDictionary<int, Row> m_rows = new SortedDictionary<int, Row>();

		/// <summary>
		/// Gets the owner document.
		/// </summary>
		/// <value>The owner document.</value>
		public SpreadSheetDocument Document
		{
			get { return document; }
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

        public Cell this[int row, int column]
        {
            get
            {
            	return this[row][column];
            }
        }

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

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.m_rows.Values.GetEnumerator();
		}
		
		IEnumerator<Row> IEnumerable<Row>.GetEnumerator()
		{
			return this.m_rows.Values.GetEnumerator();
		}
	}
}
