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
using System.Text;
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF8
{
	internal class WINDOW2 : BIFFRecord
	{
		/// <summary>
		/// BOF OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x023E;
			}
		}


		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write ( EndianStream stream )
		{
			this.WriteHeader(stream, 18);
			stream.WriteUInt16((UInt16)0x06B6); // Option flags
            stream.WriteUInt16(0); // Index to first visible row
            stream.WriteUInt16(0); // Index to first visible column
            stream.WriteUInt16((UInt16)0x0040); // Colour index of grid line colour
            stream.WriteUInt16((UInt16)0); // Not used
            stream.WriteUInt16((UInt16)0x0000); // Cached magnification factor in page break preview (in percent); 0 = Default (60%)
            stream.WriteUInt16((UInt16)0x0000); // Cached magnification factor in normal view (in percent); 0 = Default (100%)
            stream.WriteUInt32(0); // Not used
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read ( EndianStream stream )
		{
		}
	}
}

