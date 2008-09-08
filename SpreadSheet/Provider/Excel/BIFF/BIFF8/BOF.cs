using System;
using System.Collections.Generic;
using System.Text;
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Excel.BIFF.BIFF8
{
	internal static class BOF
	{
		/// <summary>
		/// BOF OPCODE
		/// </summary>
		public const int OPCODE = 0x0809;

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

		/// <summary>
		/// Writes BIFF record to the specified writer.
		/// </summary>
		/// <param name="writer">The writer.</param>
		/// <param name="type">The sheet type.</param>
		/// <param name="buildIdentifier">The build identifier.</param>
		/// <param name="historyFlags">The history flags.</param>
		public static void Write ( EndianWriter writer, SheetType type, int buildIdentifier, long historyFlags  )
		{
			RecordHeader.Write(writer, OPCODE, 14);
			writer.Write2(VERSION); // BIFF Version
			writer.Write2((int)type); // Sheet type
			writer.Write2(buildIdentifier); // Build identifier
			writer.Write2(DateTime.Today.Year); // Build year
			writer.Write4(historyFlags); // File history flags
			writer.Write4(VERSION); // Lowest Excel version that can read all records
		}
	}
}
