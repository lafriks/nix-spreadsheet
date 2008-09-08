using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

using Nix.CompoundFile;
using Nix.SpreadSheet.Provider.Excel.BIFF;
using Nix.SpreadSheet.Provider.Excel.BIFF.BIFF8;

namespace Nix.SpreadSheet.Provider
{
	public class Xls : IFileFormatProvider
	{
		#region IFileFormatProvider Members

		public void Save ( SpreadSheetDocument document, System.IO.Stream stream )
		{
			Ole2CompoundFile cf = new Ole2CompoundFile();
			
			#region Workbook stream
			MemoryStream wbs = new MemoryStream();

			EndianWriter wbsw = new LittleEndianWriter(wbs);

			BOF.Write(wbsw, BOF.SheetType.WorkBookGlobals, 0, 0);
			EOF.Write(wbsw);
			foreach(Sheet sheet in document)
			{
				BOF.Write(wbsw, BOF.SheetType.WorkSheet, 0, 0);
				EOF.Write(wbsw);
			}
			cf.Root.AddStream("Workbook", wbs);
			#endregion

			#region Summary information stream
			MemoryStream ss = new MemoryStream();
			EndianWriter ssw = new LittleEndianWriter(ss);
			cf.Root.AddStream((char)0x05 + "SummaryInformation", ss);
			#endregion

			#region Document summary information stream
			MemoryStream dss = new MemoryStream();
			EndianWriter dssw = new LittleEndianWriter(dss);
			cf.Root.AddStream((char)0x05 + "DocumentSummaryInformation", dss);
			#endregion

			//cf.Root.AddStream("Workbook", mem.ToArray());
			//cf.Root.AddStream("SummaryInformation", mem.ToArray());
			//cf.Root.AddStream("DocumentSummaryInformation", mem.ToArray());
			cf.Save(stream);
			wbs.Close();
			ss.Close();
			dss.Close();
		}

		#endregion
	}
}
