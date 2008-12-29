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

namespace Nix.SpreadSheet.Provider.Excel.BIFF.BIFF8
{
	internal class BOF : BIFFRecord
	{
		/// <summary>
		/// BOF OPCODE
		/// </summary>
		protected override int OPCODE {
			get {
				return 0x0809;
			}
		}

		/// <summary>
		/// BIFF Version
		/// </summary>
		public const int VERSION = 0x0600;

		public enum SheetType
		{
			WorkBookGlobals = 0x0005,
			VisalBasicModule = 0x0006,
			WorkSheet = 0x0010,
			Chart = 0x0020,
			MacroSheet = 0x0040,
			WorkSpace = 0x0100
		}

		[Flags]
		public enum HistoryFlag
		{
			None = 0x00000000,
			LastEditOnWindows = 0x00000001,
			LastEditOnRISC = 0x00000002,
			LastEditByBetaVersion = 0x00000004,
			EverEditedOnWindows = 0x00000008,
			EverEditedOnMacintosh = 0x00000010,
			EverEditedByBetaVersion = 0x00000020,
			EverEditedOnRISC = 0x00000100
		}

		private SheetType type = SheetType.WorkSheet;

		/// <summary>
		/// BOF record type.
		/// </summary>
		public SheetType Type {
			get { return type; }
			set { type = value; }
		}

		private int buildIdentifier = 0x0DBB;

		/// <summary>
		/// Build identifier (Default for Excel 97).
		/// </summary>
		public int BuildIdentifier {
			get { return buildIdentifier; }
			set { buildIdentifier = value; }
		}

		private int buildYear = 1996;

		/// <summary>
		/// Build year. Defaults to 1996 (Excel 97).
		/// </summary>
		public int BuildYear {
			get { return buildYear; }
			set { buildYear = value; }
		}

		private HistoryFlag historyFlags = HistoryFlag.None;

		/// <summary>
		/// History flags
		/// </summary>
		public HistoryFlag HistoryFlags {
			get { return historyFlags; }
			set { historyFlags = value; }
		}

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write ( EndianStream stream )
		{
			this.WriteHeader(stream, 16);
			stream.Write2(VERSION); // BIFF Version
			stream.Write2((int)this.Type); // Sheet type
			stream.Write2(this.BuildIdentifier); // Build identifier
			stream.Write2(this.BuildYear); // Build year
			stream.Write4((long)this.HistoryFlags); // File history flags
			stream.Write4(VERSION); // Lowest Excel version that can read all records
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read ( EndianStream stream )
		{
			int lenght = this.ReadHeader(stream);
			if ( lenght != 14 )
				throw new IndexOutOfRangeException("Invalid BOF record");
			int version = stream.Read2(); // BIFF Version
			if ( version != VERSION )
				throw new NotSupportedException("Unsupported BIFF version");
			this.Type = (SheetType)stream.Read2(); // Sheet type
			this.BuildIdentifier = stream.Read2(); // Build identifier
			this.BuildYear = stream.Read2(); // Build year
			this.HistoryFlags = (HistoryFlag)stream.Read4(); // File history flags
			long lowest_version = stream.Read4(); // Lowest Excel version that can read all records
			if ( VERSION < lowest_version )
				throw new NotSupportedException("Unsupported BIFF version");
		}
	}
}
