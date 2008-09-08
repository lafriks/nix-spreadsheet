using System;
using System.Collections.Generic;
using System.Text;
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Excel.BIFF
{
	/// <summary>
	/// BIFF record base class.
	/// </summary>
	internal static class RecordHeader
	{
		/// <summary>
		/// Writes record header to the specified writer.
		/// </summary>
		/// <param name="pWriter">The writer.</param>
		/// <param name="nRecNo">The record type.</param>
		/// <param name="nRecLen">The record length.</param>
        public static void Write (EndianWriter writer, int recordType, int recordLength)
        {
			writer.Write2(recordType);
			writer.Write2(recordLength);
        }
	}
}
