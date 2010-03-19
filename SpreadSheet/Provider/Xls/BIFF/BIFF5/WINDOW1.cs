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

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF5
{
	internal class WINDOW1 : BIFFRecord
	{
		/// <summary>
		/// BOF OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x003D;
			}
		}


		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write ( EndianStream stream )
		{
			this.WriteHeader(stream, 18);
			stream.WriteUInt16((UInt16)0x0078); // Horizontal position of the document window (in twips = 1/20 of a point)
            stream.WriteUInt16((UInt16)0x001E); // Vertical position of the document window (in twips = 1/20 of a point)
            stream.WriteUInt16((UInt16)0x559B); // Width of the document window (in twips = 1/20 of a point)
            stream.WriteUInt16((UInt16)0x32EB); // Height of the document window (in twips = 1/20 of a point)
            stream.WriteUInt16((UInt16)0x0038); // Option flags
            stream.WriteUInt16((UInt16)0x0000); // Index to active (displayed) worksheet
            stream.WriteUInt16((UInt16)0x0000); // Index of first visible tab in the worksheet tab bar
            stream.WriteUInt16((UInt16)0x0001); // Number of selected worksheets (highlighted in the worksheet tab bar)
            stream.WriteUInt16((UInt16)0x0258); // Width of worksheet tab bar (in 1/1000 of window width)
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

