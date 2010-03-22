using System;
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
	class CONTINUE : BIFFRecord
	{
		public ushort RecordLength { get; set; }

		/// <summary>
		/// CONTINUE OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x003C;
			}
		}

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write(EndianStream stream)
		{
			this.WriteHeader(stream, RecordLength);
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
