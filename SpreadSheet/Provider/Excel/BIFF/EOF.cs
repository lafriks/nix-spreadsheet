using System;
using System.Collections.Generic;
using System.Text;
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Excel.BIFF
{
	internal static class EOF
	{
		public const int OPCODE = 0x000A;
	
		public static void Write ( EndianWriter writer )
		{
			RecordHeader.Write(writer, OPCODE, 0);
		}
	}
}
