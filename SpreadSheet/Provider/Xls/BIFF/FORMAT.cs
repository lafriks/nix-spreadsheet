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
	/// FORMAT BIFF record.
	/// </summary>
	internal class FORMAT : BIFFRecord
	{
		protected override int OPCODE {
			get {
				return 0x41E;
			}
		}

		private int index = 0;
		
		public int Index {
			get { return index; }
			set { index = value; }
		}

		private string format = string.Empty;
		
		public string Format {
			get { return format; }
			set { format = value; }
		}

		public override void Write(Nix.CompoundFile.EndianStream stream)
		{
			this.WriteHeader(stream, 2 + BIFFStringHelper.GetStringByteCount(this.Format));
			stream.Write2(this.Index);
			BIFFStringHelper.WriteString(stream, this.Format);
		}

		public override void Read(Nix.CompoundFile.EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
