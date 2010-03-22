using System;
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
	class BLANK : CellBIFFRecord
	{
		/// <summary>
		/// CONTINUE OPCODE
		/// </summary>
		protected override ushort OPCODE
		{
			get
			{
				return 0x0201;
			}
		}

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write(EndianStream stream)
		{
			this.WriteHeader(stream, 6);
			stream.WriteUInt16(RowIndex);
			stream.WriteUInt16(ColIndex);
			stream.WriteUInt16(XfIndex);
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read(EndianStream stream)
		{
			throw new NotImplementedException();
		}
	}
}
