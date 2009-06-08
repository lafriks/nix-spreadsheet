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
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF5
{
	/// <summary>
	/// Description of ROW.
	/// </summary>
	internal class ROW : BIFFRecord
	{
		/// <summary>
		/// ROW OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x0208;
			}
		}

		private Row row;
		
		public Row Row {
			get { return row; }
			set { row = value; }
		}

		private int firstCellIndex;
		
		public int FirstCellIndex {
			get { return firstCellIndex; }
			set { firstCellIndex = value; }
		}

		private int lastCellIndex;
		
		public int LastCellIndex {
			get { return lastCellIndex; }
			set { lastCellIndex = value; }
		}

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write(EndianStream stream)
		{
			this.WriteHeader(stream, 16);
			stream.WriteUInt16((ushort)this.Row.RowIndex);
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read(EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
