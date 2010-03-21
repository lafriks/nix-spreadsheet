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
using System.Collections.Generic;
using Nix.CompoundFile;
using Nix.SpreadSheet.Provider.Xls.BIFF;

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF8
{
	/// <summary>
	/// SST BIFF record.
	/// </summary>
	internal class SST : BIFFRecord
	{
		public uint TotalCount { get; set; }
		public List<string> StringTable { get; set; }
		
		/// <summary>
		/// BOF OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x00FC;
			}
		}

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write ( EndianStream stream )
		{
			ushort reclen = 8;
			foreach (string str in StringTable)
				reclen += BIFFStringHelper.GetStringByteCount(str, true);
			this.WriteHeader(stream, reclen);
			stream.WriteUInt32(TotalCount); // Total string count
			stream.WriteUInt32((uint)StringTable.Count); // Unique string count
			foreach (string str in StringTable)
				BIFFStringHelper.WriteString(stream, str, true);
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read ( EndianStream stream )
		{
			throw new NotImplementedException();
		}
	}
}
