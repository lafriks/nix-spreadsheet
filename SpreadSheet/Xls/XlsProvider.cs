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
using System.IO;

using Nix.CompoundFile;
using Nix.SpreadSheet.Provider.Excel.BIFF;
using Nix.SpreadSheet.Provider.Excel.BIFF.BIFF8;

namespace Nix.SpreadSheet.Provider
{
	public class XlsProvider : IFileFormatProvider
	{
		private EndianStream activeStream = null;
		private void Write(BIFFRecord record)
		{
			if ( activeStream == null )
				throw new ArgumentNullException("activeStream");
			record.Write(activeStream);
		}

		private List<Font> fontTable = new List<Font>();

		public int FindFontIndex(Font font)
		{
			for (int i = 0; i < fontTable.Count; i++)
			{
				if ( fontTable[i].Equals(font) )
				{
					if ( i >= 4 )
						return i + 1;
					else
						return i;
				}
			}
			return -1;
		}
		
		protected void BuildFontTable( SpreadSheetDocument document )
		{
			foreach(Sheet sheet in document)
				foreach(Cell cell in sheet)
					if ( FindFontIndex(cell.Style.Font) == -1 )
						fontTable.Add(cell.Style.Font);
			// There should be at least 5 fonts in the table
			for ( int i = fontTable.Count; i < 5; i++ )
				fontTable.Add(Font.Default);
		}
		#region IFileFormatProvider Members

		public void Save ( SpreadSheetDocument document, System.IO.Stream stream )
		{
			Ole2CompoundFile cf = new Ole2CompoundFile();

			this.BuildFontTable(document);
			
			#region Workbook stream
			MemoryStream wbs = new MemoryStream();

			this.activeStream = new LittleEndianStream(wbs);

			// Workbook globals
			this.Write(new BOF() { Type = BOF.SheetType.WorkBookGlobals });
			// Font table
			foreach ( Font font in this.fontTable )
				this.Write(new FONT() { Font = font });
			this.Write(new EOF());

			// Sheets
			foreach(Sheet sheet in document)
			{
				this.Write(new BOF() { Type = BOF.SheetType.WorkSheet});
				this.Write(new EOF());
			}

			cf.Root.AddStream("Workbook", wbs);
			#endregion

			#region Summary information stream
			MemoryStream ss = new MemoryStream();
			this.activeStream = new LittleEndianStream(ss);
			cf.Root.AddStream((char)0x05 + "SummaryInformation", ss);
			#endregion

			#region Document summary information stream
			MemoryStream dss = new MemoryStream();
			this.activeStream = new LittleEndianStream(dss);
			cf.Root.AddStream((char)0x05 + "DocumentSummaryInformation", dss);
			#endregion

			cf.Save(stream);
			wbs.Close();
			ss.Close();
			dss.Close();
		}

		#endregion
	}
}
