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
using System.Drawing;

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
			if ( ! (this.Style is CellStyle) )
				return 0;
			byte used_bits = 0;
			if ( ((CellStyle)this.Style).IsModifiedFormat() )
				used_bits |= 0x01;
			if ( ((CellStyle)this.Style).IsModifiedFont() )
				used_bits |= 0x02;
			if ( ((CellStyle)this.Style).IsModifiedHorizontalAlignment() || ((CellStyle)this.Style).IsModifiedVerticalAlignment()
			    || ((CellStyle)this.Style).IsModifiedWrapTextAtRightBorder() || ((CellStyle)this.Style).IsModifiedJustifyTextAtLastLine()
			    || ((CellStyle)this.Style).IsModifiedWrapText() || ((CellStyle)this.Style).IsModifiedRotation()
			    || ((CellStyle)this.Style).IsModifiedShrinkToFit() || ((CellStyle)this.Style).IsModifiedTextDirection()
			    || ((CellStyle)this.Style).IsModifiedIndentLevel() )
				used_bits |= 0x04;
			if ( ((CellStyle)this.Style).IsModifiedTopBorderLineStyle() || ((CellStyle)this.Style).IsModifiedTopBorderLineColor()
			    || ((CellStyle)this.Style).IsModifiedLeftBorderLineStyle() || ((CellStyle)this.Style).IsModifiedLeftBorderLineColor()
			    || ((CellStyle)this.Style).IsModifiedRightBorderLineStyle() || ((CellStyle)this.Style).IsModifiedRightBorderLineColor()
			    || ((CellStyle)this.Style).IsModifiedBottomBorderLineStyle() || ((CellStyle)this.Style).IsModifiedBottomBorderLineColor() )
				used_bits |= 0x08;
			if ( ((CellStyle)this.Style).IsModifiedBackgroundColor() || ((CellStyle)this.Style).IsModifiedBackgroundPattern()
			    || ((CellStyle)this.Style).IsModifiedBackgroundPatternColor() )
				used_bits |= 0x10;
			if ( ((CellStyle)this.Style).IsModifiedCellLocked() || ((CellStyle)this.Style).IsModifiedHiddenFormula() )
				used_bits |= 0x20;
			return used_bits;
		}

		private byte GetTextDirection()
		{
			switch(this.Style.TextDirection)
			{
				default:
				case TextDirection.Automatic:
					return 0;
				case TextDirection.LeftToRight:
					return 1;
				case TextDirection.RightToLeft:
					return 2;
			}
		}

		private byte GetLineStyle(BorderLineStyle style)
		{
			switch (style)
			{
				default:
				case BorderLineStyle.None:
					return 0;
				case BorderLineStyle.Thin:
					return 1;
				case BorderLineStyle.Medium:
					return 2;
				case BorderLineStyle.Dashed:
					return 3;
				case BorderLineStyle.Dotted:
					return 4;
				case BorderLineStyle.Thick:
					return 5;
				case BorderLineStyle.Double:
					return 6;
				case BorderLineStyle.Hair:
					return 7;
				case BorderLineStyle.MediumDashed:
					return 8;
				case BorderLineStyle.ThinDashDotted:
					return 9;
				case BorderLineStyle.MediumDashDotted:
					return 10;
				case BorderLineStyle.ThinDashDotDotted:
					return 11;
				case BorderLineStyle.MediumDashDotDotted:
					return 12;
				case BorderLineStyle.SlantedMediumDashDotted:
					return 13;
			}
		}

		private byte GetBgPattern()
		{
			switch (this.Style.BackgroundPattern)
			{
				default:
				case CellBackgroundPattern.None:
					return 0;
				case CellBackgroundPattern.Fill:
					return 1;
				case CellBackgroundPattern.HorizontalLines:
					return 11;
				case CellBackgroundPattern.VerticalLines:
					return 12;
			}
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
			// Indent level, shrink, text direction (1)
			byte ind_byte = this.Style.IndentLevel;
			if ( this.Style.ShrinkToFit )
				ind_byte |= 0x10;
			ind_byte |= (byte)(this.GetTextDirection() << 6);
			stream.WriteByte(ind_byte);
			// Flags for used atribute groups
			// TODO: For now parent style elements are all valid
			stream.WriteByte((byte)((this.ParentStyleIndex.HasValue ? this.GetUsedAtributes() : 0) << 2));
			// Borders (4)
			uint bord_bits = (uint)(this.GetLineStyle(this.Style.LeftBorderLineStyle)
				| (this.GetLineStyle(this.Style.RightBorderLineStyle) << 4)
				| (this.GetLineStyle(this.Style.TopBorderLineStyle) << 8)
				| (this.GetLineStyle(this.Style.BottomBorderLineStyle) << 12));
			bord_bits |= ((uint)ColorTranslator.ToOle(this.Style.LeftBorderLineColor) << 16)
				| ((uint)ColorTranslator.ToOle(this.Style.RightBorderLineColor) << 23);
			// TODO: Diagonal lines to show
			stream.WriteUInt32(bord_bits);
			// Background (4)
			uint back_bits = (uint)ColorTranslator.ToOle(this.Style.TopBorderLineColor)
				| ((uint)ColorTranslator.ToOle(this.Style.BottomBorderLineColor) << 7);
			// TODO: Diagonal line color and style
			back_bits |= ((uint)this.GetBgPattern() << 26);
			stream.WriteUInt32(back_bits);
			// Background color (2);
			ushort bgcol_bits = (ushort)((ushort)ColorTranslator.ToOle(this.Style.BackgroundPatternColor)
			                             | (ushort)(ColorTranslator.ToOle(this.Style.BackgroundColor) << 7));
			stream.WriteUInt16(bgcol_bits);
		}

		public override void Read(Nix.CompoundFile.EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
