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
		
		private Style style = null;

		public Style Style
		{
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
		
		private byte GetHorizontalAlignment(CellHorizontalAlignment align)
		{
			switch ( align )
			{
				default:
				case CellHorizontalAlignment.General:
					return 0;
				case CellHorizontalAlignment.Left:
					return 1;
				case CellHorizontalAlignment.Centred:
					return 2;
				case CellHorizontalAlignment.Right:
					return 3;
				case CellHorizontalAlignment.Filled:
					return 4;
				case CellHorizontalAlignment.Justified:
					return 5;
				case CellHorizontalAlignment.CentredAcrossSelection:
					return 6;
				case CellHorizontalAlignment.Distributed:
					return 7;
			}
		}

		private byte GetVerticalAlignment ( CellVerticalAlignment align )
		{
			switch ( align )
			{
				default:
				case CellVerticalAlignment.Top:
					return 0;
				case CellVerticalAlignment.Centred:
					return 1;
				case CellVerticalAlignment.Bottom:
					return 2;
				case CellVerticalAlignment.Justified:
					return 3;
				case CellVerticalAlignment.Distributed:
					return 4;
			}
		}

		private byte GetUsedAtributes()
		{
			byte used_bits = 0;
			if ( this.Style.IsModifiedFormat() )
				used_bits |= 0x01;
			if ( this.Style.IsModifiedFont() )
				used_bits |= 0x02;
			// TODO: Indent level, shrink content, text direction, orientation
			if ( this.Style.IsModifiedHorizontalAlignment() || this.Style.IsModifiedVerticalAlignment()
					|| this.Style.IsModifiedWrapTextAtRightBorder() || this.Style.IsModifiedJustifyTextAtLastLine()
					|| this.Style.IsModifiedWrapText() || this.Style.IsModifiedRotation() )
				used_bits |= 0x04;
			if ( this.Style.IsModifiedTopBorderLineStyle() || this.Style.IsModifiedTopBorderLineColor()
					|| this.Style.IsModifiedLeftBorderLineStyle() || this.Style.IsModifiedLeftBorderLineColor()
					|| this.Style.IsModifiedRightBorderLineStyle() || this.Style.IsModifiedRightBorderLineColor()
					|| this.Style.IsModifiedBottomBorderLineStyle() || this.Style.IsModifiedBottomBorderLineColor() )
				used_bits |= 0x08;
			if ( this.Style.IsModifiedBackgroundColor() || this.Style.IsModifiedBackgroundPattern()
					|| this.Style.IsModifiedBackgroundPatternColor() )
				used_bits |= 0x10;
			if ( this.Style.IsModifiedCellLocked() || this.Style.IsModifiedHiddenFormula() )
				used_bits |= 0x20;
			return used_bits;
		}

		public override void Write(Nix.CompoundFile.EndianStream stream)
		{
			this.WriteHeader(stream, 20);
			// Font (2)
			stream.WriteUInt16(this.FontIndex);
			// Format (2)
			stream.WriteUInt16(this.FormatIndex);
			// Cell protection and parent style (2)
			ushort prot_bit = 0; // (2 - 0)
			if ( this.Style.CellLocked )
				prot_bit &= 0x01;
			if ( this.Style.HiddenFormula )
				prot_bit &= 0x02;
			if ( ! this.ParentStyleIndex.HasValue )
				prot_bit &= 0x04;
			prot_bit = (ushort)(((this.ParentStyleIndex.HasValue ? this.ParentStyleIndex.Value : 0xFFF) << 4)
								& prot_bit); // Parent index (15 - 4)
			stream.WriteUInt16(prot_bit);
			// Alignment and text break (1)
		    byte align_bit = this.GetHorizontalAlignment(this.Style.HorizontalAlignment);
		    if ( this.Style.WrapTextAtRightBorder )
				align_bit |= 0x08;
			align_bit |= (byte)(this.GetVerticalAlignment(this.Style.VerticalAlignment) << 4);
			stream.WriteByte(align_bit);
			// Rotation angle (1)
			stream.WriteByte(this.Style.Rotation);
			// TODO: Indent level, shrink, text direction (1)
			stream.WriteByte(0);
			// Flags for used atribute groups
			// TODO: For now parent style elements are all valid
			stream.WriteByte((byte)((this.ParentStyleIndex.HasValue ? this.GetUsedAtributes() : 0) << 2));
		}

		public override void Read(Nix.CompoundFile.EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
