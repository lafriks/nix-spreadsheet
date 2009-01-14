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

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
	/// <summary>
	/// Description of XF.
	/// </summary>
	internal class XF : BIFFRecord
	{
		protected override ushort OPCODE {
			get {
				return 0xE0;
			}
		}
		
		private CellStyle style = null;
		
		public CellStyle Style {
			get { return style; }
			set { style = value; }
		}

		private ushort fontIndex = 0;
		
		public ushort FontIndex {
			get { return fontIndex; }
			set { fontIndex = value; }
		}

		private ushort formatIndex = 0;
		
		public ushort FormatIndex {
			get { return formatIndex; }
			set { formatIndex = value; }
		}

		private ushort? parentStyleIndex = null;
		
		public ushort? ParentStyleIndex {
			get { return parentStyleIndex; }
			set { parentStyleIndex = value; }
		}

		public override void Write(Nix.CompoundFile.EndianStream stream)
		{
			this.WriteHeader(stream, 20);
			stream.WriteUInt16(this.FontIndex);
			stream.WriteUInt16(this.FormatIndex);
			ushort prot_bit = 0;
			if ( this.Style.CellLocked )
				prot_bit &= 0x01;
			if ( this.Style.HiddenFormula )
				prot_bit &= 0x02;
			if ( ! this.ParentStyleIndex.HasValue )
				prot_bit &= 0x04;
			prot_bit = (ushort)(((this.ParentStyleIndex.HasValue ? this.ParentStyleIndex.Value : 0xFFF) << 3)
			                    & prot_bit);
			stream.WriteUInt16(prot_bit);
		}

		public override void Read(Nix.CompoundFile.EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
