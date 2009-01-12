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
		protected override int OPCODE {
			get {
				return 0xE0;
			}
		}
		
		private CellStyle style = null;
		
		public CellStyle Style {
			get { return style; }
			set { style = value; }
		}

		private int fontIndex = 0;
		
		public int FontIndex {
			get { return fontIndex; }
			set { fontIndex = value; }
		}

		private int formatIndex = 0;
		
		public int FormatIndex {
			get { return formatIndex; }
			set { formatIndex = value; }
		}

		private int? parentStyleIndex = null;
		
		public int? ParentStyleIndex {
			get { return parentStyleIndex; }
			set { parentStyleIndex = value; }
		}

		public override void Write(Nix.CompoundFile.EndianStream stream)
		{
			this.WriteHeader(stream, 20);
			stream.Write2(this.FontIndex);
			stream.Write2(this.FormatIndex);
			int prot_bit = 0;
			if ( this.Style.CellLocked )
				prot_bit &= 0x01;
			if ( this.Style.HiddenFormula )
				prot_bit &= 0x02;
			if ( ! this.ParentStyleIndex.HasValue )
				prot_bit &= 0x04;
			prot_bit = ((this.ParentStyleIndex.HasValue ? this.ParentStyleIndex.Value : 0xFFF) << 3)
							& prot_bit;
			stream.Write2(prot_bit);
		}

		public override void Read(Nix.CompoundFile.EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
