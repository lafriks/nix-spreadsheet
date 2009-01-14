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

using Nix.SpreadSheet;
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
	/// <summary>
	/// Description of String.
	/// </summary>
	public static class BIFFStringHelper
	{
		private static byte GetGRBIT(string text, StringFormating[] formating, bool stringLengthInt)
		{
			byte grbit = 0;
			foreach (char c in text)
				if ( c > 255 )
					grbit |= 0x01;
			if ( formating != null )
				grbit |= 0x08;
			return grbit;
		}

		public static ushort GetStringByteCount(string text, StringFormating[] formating, bool stringLengthInt)
		{
			byte grbit = GetGRBIT(text, formating, stringLengthInt);
			// String length byte(s) + grbit byte
			ushort length = (ushort)(stringLengthInt ? 3 : 2);
			 // Formating run count bytes
			if ( (grbit & 0x08) == 0x08 )
				length += 2;
			if ( (grbit & 0x01) == 0x01 )
				length += (ushort)(text.Length * 2);
			else
				length += (ushort)text.Length;
			// TODO: formating
			return length;
		}

		public static ushort GetStringByteCount(string text, bool stringLengthInt)
		{
			return GetStringByteCount(text, null, stringLengthInt);
		}

		public static ushort GetStringByteCount(string text)
		{
			return GetStringByteCount(text, null, true);
		}

		public static void WriteString(EndianStream stream, string text, StringFormating[] formating, bool stringLengthInt)
		{
			// Text length
			if ( stringLengthInt )
				stream.WriteUInt16((ushort)text.Length);
			else
				stream.WriteByte((byte)text.Length);

			byte grbit = GetGRBIT(text, formating, stringLengthInt);

			stream.WriteByte(grbit); // String options

			if ( (grbit & 0x08) == 0x08 )
				stream.WriteUInt16((ushort)formating.Length); // Formating run count

			if ( (grbit & 0x01) == 0x01 )
			{
				// Write uncompressed string
				foreach ( char c in text )
					stream.WriteUInt16((ushort)c);
			}
			else
			{
				// Write compressed string
				foreach ( char c in text )
					stream.WriteByte((byte)c);
			}

			// TODO: Write formating
		}

		public static void WriteString(EndianStream stream, string text, bool stringLengthInt)
		{
			WriteString(stream, text, null, stringLengthInt);
		}

		public static void WriteString(EndianStream stream, string text)
		{
			WriteString(stream, text, true);
		}
	}
}
