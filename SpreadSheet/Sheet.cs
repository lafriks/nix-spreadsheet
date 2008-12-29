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
	public class Sheet : IEnumerable
	{
		private SpreadSheetDocument document;

		public const int MaxRows = 65536;
        public const int MaxColumns = 256;

        private Hashtable m_table = new Hashtable();
	    

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

        public Cell this[int row, int column]
        {
            get
            {
                if (row >= MaxRows || column >= MaxColumns)
                    throw new ArgumentOutOfRangeException();
                string key = Utils.CellName(row, column);
                if (this.m_table.ContainsKey(key))
                {
                    return (Cell)this.m_table[key];
                }
                else
                {
                    Cell nc = new Cell(row, column);
                    this.m_table.Add(key, nc);
                    return nc;
                }
            }
        }

        public Cell this[string name]
        {
            get
            {
                if (this.m_table.ContainsKey(name))
                {
                    return (Cell)this.m_table[name];
                }
                else
                {
                    Cell nc = new Cell(name);
                    if (nc.RowIndex >= MaxRows || nc.ColumnIndex >= MaxColumns)
                        throw new ArgumentOutOfRangeException();
                    this.m_table.Add(nc.Name, nc);
                    return nc;
                }
            }
        }

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.m_table.Values.GetEnumerator();
		}
	}
}
