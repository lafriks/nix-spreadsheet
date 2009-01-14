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
using Nix.SpreadSheet.Provider.Xls.BIFF;
using Nix.SpreadSheet.Provider.Xls.BIFF.BIFF8;

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
		
		protected void BuildFontTable ( SpreadSheetDocument document )
		{
			foreach(Sheet sheet in document)
				foreach(Cell cell in sheet)
					if ( FindFontIndex(cell.Formatting.Font) == -1 )
						fontTable.Add(cell.Formatting.Font);
			// There should be at least 5 fonts in the table
			for ( int i = fontTable.Count; i < 5; i++ )
				fontTable.Add(Font.Default);
		}

		private Dictionary<ushort, string> formatTable = new Dictionary<ushort,string>();

		private List<ushort> formatsToWrite = new List<ushort>();

		private ushort formatTableSeq = 0;

		protected void BuildFormatTable ( SpreadSheetDocument document )
		{
			// Add default formats
			this.formatTable.Add(0x00, "General");
			this.formatTable.Add(0x01, "0");
			this.formatTable.Add(0x02, "0.00");
			this.formatTable.Add(0x03, "#,##0");
			this.formatTable.Add(0x04, "#,##0.00");
			if ( document.Locale.TwoLetterISORegionName == "LV")
			{
				this.formatTable.Add(0x05, "#,##0\\ \"Ls\";\\-#,##0\\ \"Ls\"");
				this.formatsToWrite.Add(0x05);
				this.formatTable.Add(0x06, "#,##0\\ \"Ls\";[Red]\\-#,##0\\ \"Ls\"");
				this.formatsToWrite.Add(0x06);
				this.formatTable.Add(0x07, "#,##0.00\\ \"Ls\";\\-#,##0.00\\ \"Ls\"");
				this.formatsToWrite.Add(0x07);
				this.formatTable.Add(0x08, "#,##0.00\\ \"Ls\";[Red]\\-#,##0.00\\ \"Ls\"");
				this.formatsToWrite.Add(0x08);
			}
			else
			{
				this.formatTable.Add(0x05, "($#,##0_);($#,##0)");
				this.formatTable.Add(0x06, "($#,##0_);[Red]($#,##0)");
				this.formatTable.Add(0x07, "($#,##0.00_);($#,##0.00)");
				this.formatTable.Add(0x08, "($#,##0.00_);[Red]($#,##0.00)");
			}
			this.formatTable.Add(0x09, "0%");
			this.formatTable.Add(0x0a, "0.00%");
			this.formatTable.Add(0x0b, "0.00E+00");
			this.formatTable.Add(0x0c, "# ?/?");
			this.formatTable.Add(0x0d, "# ??/??");
			this.formatTable.Add(0x0e, "m/d/yy");
			this.formatTable.Add(0x0f, "d-mmm-yy");
			this.formatTable.Add(0x10, "d-mmm");
			this.formatTable.Add(0x11, "mmm-yy");
			this.formatTable.Add(0x12, "h:mm AM/PM");
			this.formatTable.Add(0x13, "h:mm:ss AM/PM");
			this.formatTable.Add(0x14, "h:mm");
			this.formatTable.Add(0x15, "h:mm:ss");
			this.formatTable.Add(0x16, "m/d/yy h:mm");
			this.formatTable.Add(0x25, "(#,##0_);(#,##0)");
			this.formatTable.Add(0x26, "(#,##0_);[Red](#,##0)");
			this.formatTable.Add(0x27, "(#,##0.00_);(#,##0.00)");
			this.formatTable.Add(0x28, "(#,##0.00_);[Red](#,##0.00)");
			if ( document.Locale.TwoLetterISORegionName == "LV")
			{
				this.formatTable.Add(0x29, "_-* #,##0\\ _L_s_-;\\-* #,##0\\ _L_s_-;_-* \"-\"\\ _L_s_-;_-@_-");
				this.formatsToWrite.Add(0x29);
				this.formatTable.Add(0x2a, "_-* #,##0\\ \"Ls\"_-;\\-* #,##0\\ \"Ls\"_-;_-* \"-\"\\ \"Ls\"_-;_-@_-");
				this.formatsToWrite.Add(0x2a);
				this.formatTable.Add(0x2b, "_-* #,##0.00\\ _L_s_-;\\-* #,##0.00\\ _L_s_-;_-* \"-\"??\\ _L_s_-;_-@_-");
				this.formatsToWrite.Add(0x2b);
				this.formatTable.Add(0x2c, "_-* #,##0.00\\ \"Ls\"_-;\\-* #,##0.00\\ \"Ls\"_-;_-* \"-\"??\\ \"Ls\"_-;_-@_-");
				this.formatsToWrite.Add(0x2c);
			}
			else
			{
				this.formatTable.Add(0x29, "_(* #,##0_);_(* (#,##0);_(* \" - \"_);_(@_)");
				this.formatTable.Add(0x2a, "_($* #,##0_);_($* (#,##0);_($* \" - \"_);_(@_)");
				this.formatTable.Add(0x2b, "_(* #,##0.00_);_(* (#,##0.00);_(* \" - \"??_);_(@_)");
				this.formatTable.Add(0x2c, "_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)");
			}
			this.formatTable.Add(0x2d, "mm:ss");
			this.formatTable.Add(0x2e, "[h]:mm:ss");
			this.formatTable.Add(0x2f, "mm:ss.0");
			this.formatTable.Add(0x30, "##0.0E+0");
			this.formatTable.Add(0x31, "@");

			this.formatTableSeq = 0x32;

			// Add user defined formats if they are not in default format table
			foreach ( Sheet sheet in document )
			{
				foreach ( Cell cell in sheet )
				{
					if ( ! this.formatTable.ContainsValue(cell.Formatting.Format) )
					{
						this.formatTable.Add(this.formatTableSeq, cell.Formatting.Format);
						this.formatsToWrite.Add(this.formatTableSeq);
						this.formatTableSeq++;
					}
				}
			}
		}

		#region IFileFormatProvider Members

		public void Save ( SpreadSheetDocument document, System.IO.Stream stream )
		{
			Ole2CompoundFile cf = new Ole2CompoundFile();

			this.BuildFontTable(document);
			this.BuildFormatTable(document);
			
			#region Workbook stream
			MemoryStream wbs = new MemoryStream();

			this.activeStream = cf.CreateEndianStream(wbs);

			// Workbook globals
			this.Write(new BOF() { Type = BOF.SheetType.WorkBookGlobals });
			// Font table
			foreach ( Font font in this.fontTable )
				this.Write(new FONT() { Font = font });
			// Format table
			foreach ( ushort fi in this.formatsToWrite )
				this.Write(new FORMAT() { Index = fi, Format = this.formatTable[fi] });
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
			this.activeStream = cf.CreateEndianStream(ss);
			cf.Root.AddStream((char)0x05 + "SummaryInformation", ss);
			#endregion

			#region Document summary information stream
			MemoryStream dss = new MemoryStream();
			this.activeStream = cf.CreateEndianStream(dss);
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
