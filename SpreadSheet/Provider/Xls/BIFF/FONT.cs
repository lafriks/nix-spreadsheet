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
	/// Description of FONT.
	/// </summary>
	internal class FONT : BIFFRecord
	{
		protected override ushort OPCODE {
			get {
				return 0x231;
			}
		}

		private Font font = null;
		
		public Font Font {
			get { return font; }
			set { font = value; }
		}

		private ushort ScriptPositionToInt(ScriptPosition sp)
		{
			switch (sp)
			{
				default:
				case ScriptPosition.Normal:
					return 0x00;
				case ScriptPosition.Superscript:
					return 0x01;
				case ScriptPosition.Subscript:
					return 0x02;
			}
		}

		private byte UnderlineStyleToByte(UnderlineStyle us)
		{
			switch (us)
			{
				default:
				case UnderlineStyle.None:
					return 0x00;
				case UnderlineStyle.Single:
					return 0x01;
				case UnderlineStyle.Double:
					return 0x02;
				case UnderlineStyle.SingleAccounting:
					return 0x21;
				case UnderlineStyle.DoubleAccounting:
					return 0x22;
			}
		}

		private byte FontFaceToByte(FontFace ff)
		{
			switch (ff)
			{
				default:
				case FontFace.None:
					return 0;
				case FontFace.Decorative:
					return 80;
				case FontFace.Modern:
					return 48;
				case FontFace.Roman:
					return 16;
				case FontFace.Script:
					return 64;
				case FontFace.Swiss:
					return 32;
			}
		}

		private byte CharSetToByte(CharSet cs)
		{
			switch(cs)
			{
				default:
				case CharSet.Ansi:
					return 0;
				case CharSet.Arabic:
					return 178;
				case CharSet.Baltic:
					return 186;
				case CharSet.ChineseBig5:
					return 136;
				case CharSet.Default:
					return 1;
				case CharSet.EastEurope:
					return 238;
				case CharSet.GB2312:
					return 134;
				case CharSet.Greek:
					return 161;
				case CharSet.Hangeul:
					return 129;
				case CharSet.Hebrew:
					return 177;
				case CharSet.Johab:
					return 130;
				case CharSet.MAC:
					return 77;
				case CharSet.OEM:
					return 255;
				case CharSet.Russian:
					return 204;
				case CharSet.Shiftjis:
					return 128;
				case CharSet.Symbol:
					return 2;
				case CharSet.Thai:
					return 222;
				case CharSet.Turkish:
					return 162;
			}
		}

		public override void Write(Nix.CompoundFile.EndianStream stream)
		{
			this.WriteHeader(stream, (ushort)(14 + BIFFStringHelper.GetStringByteCount(this.Font.Name, false)));
			stream.WriteUInt16(this.Font.Size);
			ushort grbit = 0;
			if ( this.Font.Italic )
				grbit |= 0x02;
			if ( this.Font.Strikeout )
				grbit |= 0x08;
			// TODO: Outline and Shadow grbit
			stream.WriteUInt16(grbit); // Font options
			stream.WriteUInt16((ushort)ColorTranslator.ToOle(this.Font.Color)); // Color index
			stream.WriteUInt16(this.Font.Weight); // Font weight
			stream.WriteUInt16(ScriptPositionToInt(this.Font.ScriptPosition)); // Script position
			stream.WriteByte(UnderlineStyleToByte(this.Font.UnderlineStyle)); // Underline style
			stream.WriteByte(FontFaceToByte(this.Font.FontFace)); // Font face
			stream.WriteByte(CharSetToByte(this.Font.CharSet)); // Font charset
			stream.WriteByte(0); // Reserved
			BIFFStringHelper.WriteString(stream, this.Font.Name, false);
		}

		public override void Read(Nix.CompoundFile.EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
