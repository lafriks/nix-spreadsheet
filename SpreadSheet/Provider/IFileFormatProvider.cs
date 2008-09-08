using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Nix.SpreadSheet.Provider
{
	public interface IFileFormatProvider
	{
		void Save(SpreadSheetDocument document, Stream stream);
	}
}
