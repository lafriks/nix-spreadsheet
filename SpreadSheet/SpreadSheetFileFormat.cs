using System;
using System.Collections.Generic;
using System.Text;

namespace Nix.SpreadSheet
{
	public enum SpreadSheetFileFormat
	{
		/// <summary>
		/// Microsoft compound binary file format.
		/// </summary>
		ExcelBinary,
		/// <summary>
		/// Microsoft Open XML document format.
		/// </summary>
		OpenXml,
		/// <summary>
		/// OASIS OpenDocument, ISO/IEC 26300 file format
		/// </summary>
		OpenDocument
	}
}
