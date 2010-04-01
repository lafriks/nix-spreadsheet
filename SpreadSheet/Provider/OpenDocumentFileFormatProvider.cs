using System;
using System.Collections.Generic;
using System.Text;
using Nix.SpreadSheet.Provider.Zip;

namespace Nix.SpreadSheet.Provider
{
	public class OpenDocumentFileFormatProvider : IFileFormatProvider
	{
		public void Save(SpreadSheetDocument document, System.IO.Stream stream)
		{
			ZipStream zip = new ZipStream(stream);
			zip.AddFileWithStringContent("mimetype", "application/vnd.oasis.opendocument.spreadsheet");
		}
	}
}
