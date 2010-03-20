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

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
	internal class PALETTE : BIFFRecord
	{
		/// <summary>
		/// EOF record OPCODE.
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x0092;
			}
		}
        
		private ColourPalette palette = null;

		public ColourPalette Palette
        {
        	set { palette = value; }
        }
		
		public override void Write ( EndianStream stream )
		{
			if (palette != null && palette.Count > 0)
			{
				this.WriteHeader(stream, (ushort)(2 + 4* palette.Count));
				stream.WriteUInt16(palette.Count);
				foreach (uint color in palette)
				{
					stream.WriteUInt32(color);
				}
			}
		}

		public override void Read(EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
