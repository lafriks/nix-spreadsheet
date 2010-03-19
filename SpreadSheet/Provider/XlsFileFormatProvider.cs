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

using Nix.SpreadSheet.Provider.Xls.BIFF.BIFF5;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Nix.CompoundFile;
using Nix.SpreadSheet.Provider.Xls.BIFF;
using Nix.SpreadSheet.Provider.Xls.BIFF.BIFF8;

namespace Nix.SpreadSheet.Provider
{
	public class XlsFileFormatProvider : IFileFormatProvider
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
				foreach(Row row in sheet)
					foreach(Cell cell in row)
						if ( FindFontIndex(cell.Formatting.Font) == -1 )
						fontTable.Add(cell.Formatting.Font);
			// There should be at least 5 fonts in the table
			for ( int i = fontTable.Count; i < 5; i++ )
				fontTable.Add(Font.Default);
		}

		private Dictionary<ushort, string> formatTable = new Dictionary<ushort,string>();

		private List<ushort> formatsToWrite = new List<ushort>();

		public ushort FindFormatIndex(string format)
		{
			foreach ( ushort idx in formatTable.Keys )
				if ( formatTable[idx] == format )
					return idx;
			throw new IndexOutOfRangeException("Format not found");
		}

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
				foreach ( Row row in sheet )
				{
					foreach ( Cell cell in row )
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
		}
		

		private List<Style> styleTable = new List<Style>();

		public int FindStyleIndex(Style style)
		{
			for (int i = 0; i < styleTable.Count; i++)
			{
				if ( styleTable[i].Equals(style) )
					return i + 16;
			}
			return -1;
		}

		protected void BuildStyleTable(SpreadSheetDocument document)
		{
			// Find all parent styles
			foreach ( Sheet sheet in document )
			{
				foreach ( Row row in sheet )
				{
					foreach ( Cell cell in row )
					{
						if ( FindStyleIndex(cell.Formatting.Parent) == -1 )
						{
							this.styleTable.Add(cell.Formatting.Parent);
						}
						if ( FindStyleIndex(cell.Formatting) == -1 )
						{
							this.styleTable.Add(cell.Formatting);
						}
					}
				}
				foreach (Column col in sheet.Columns)
				{
					if ( FindStyleIndex(col.Formatting.Parent) == -1 )
					{
						this.styleTable.Add(col.Formatting.Parent);
					}
					if ( FindStyleIndex(col.Formatting) == -1 )
					{
						this.styleTable.Add(col.Formatting);
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
			this.BuildStyleTable(document);
			
			#region Workbook stream
			MemoryStream wbs = new MemoryStream();
            EndianStream mainStream = cf.CreateEndianStream(wbs);

            this.activeStream = mainStream;

            foreach (Style s in this.styleTable)
            {
                ushort format = FindFormatIndex(s.Format);
                if (!formatsToWrite.Contains(format))
                {
                    formatsToWrite.Add(format);
                }
            }

			// Workbook globals
			this.Write(new BOF() { Type = BOF.SheetType.WorkBookGlobals });
			// WINDOW1
			this.Write(new WINDOW1());
			// Font table
            foreach (Font font in this.fontTable)
                this.Write(new FONT() { Font = font });
			// Format table
			foreach ( ushort fi in this.formatsToWrite )
				this.Write(new FORMAT() { Index = fi, Format = this.formatTable[fi] });
            // Default XFs
            for (int j = 0; j < 15; j++)
            {
                XF r = new XF(){Style = this.styleTable[0],
                                FontIndex = (ushort)FindFontIndex(this.styleTable[0].Font),
                                FormatIndex = FindFormatIndex(this.styleTable[0].Format)};
                this.Write(r);

                // Default Cell XF
                if (j == 14)
                {
                    r.ParentStyleIndex = 0;
                    this.Write(r);
                }
            }
            // Style table
			foreach ( Style s in this.styleTable )
			{
				XF r = new XF() { Style = s, FontIndex = (ushort)FindFontIndex(s.Font), FormatIndex = FindFormatIndex(s.Format) };
				if ( s is CellStyle )
                    r.ParentStyleIndex = (ushort?)FindStyleIndex(((CellStyle)s).Parent);
				this.Write(r);
			}


            List<MemoryStream> sheetStreams = new List<MemoryStream>();
            uint sheetNamesLength = 0;

			// Sheets
			foreach(Sheet sheet in document)
			{
                MemoryStream tmpStream = new MemoryStream();
                activeStream = cf.CreateEndianStream(tmpStream);
                sheetStreams.Add(tmpStream);
                sheetNamesLength += BIFFStringHelper.GetStringByteCount(sheet.Name, false);

				this.Write(new BOF() { Type = BOF.SheetType.WorkSheet});
				
				foreach (Column col in sheet.Columns)
				{
					this.Write(new COLUMN() {Index = (ushort)col.ColumnIndex, Width = (ushort)col.Width, XfIndex = (ushort)FindStyleIndex(col.Formatting)});
				}
				
				this.Write(new DIMENSION() { FirstCol = (ushort)sheet.FirstColumn, FirstRow = (uint)sheet.FirstRow,
				           					 LastCol = (ushort)(sheet.LastColumn + 1), LastRow = (uint)(sheet.LastRow + 1) });
				foreach ( Row row in sheet )
				{
					this.Write( new ROW() { Row = row } );
				}
                foreach (Row row in sheet)
                {
                    foreach (Cell cell in row)
                    {
                        if (cell.Value is int || cell.Value is float || cell.Value is double ||
                            cell.Value is decimal || cell.Value is long)
                        {
                            this.Write(new CellNumber(){ColIndex = (ushort)cell.ColumnIndex,
                                                        RowIndex = (ushort)cell.RowIndex,
                                                        XfIndex = (ushort)FindStyleIndex(cell.Formatting),
                                                        Value = Convert.ToDouble(cell.Value)});
                        }
                    }
                }
				this.Write(new WINDOW2());
				this.Write(new EOF());
			}

            uint settingsLength = (uint)(document.SheetCount * 10 // Sheet base data length (without sheet name)
                                         + sheetNamesLength // Sheet name length in unicode
                                         + 4); // EOF
            
            this.activeStream = mainStream;
            uint currentPosition = (uint)wbs.Length + settingsLength;
            int i = 0;
            // Sheets' headers
            foreach(Sheet sheet in document)
            {
                this.Write(new SHEET() {Name = sheet.Name, Position = currentPosition});
                currentPosition += (uint)sheetStreams[i].Length;
                i++;
            }
            this.Write(new EOF());

            foreach (MemoryStream str in sheetStreams)
            {
                wbs.Write(str.GetBuffer(), 0, (int)str.Length);
            }

			cf.Root.AddStream("Workbook", wbs);

			#endregion
/*
			#region Summary information stream
			MemoryStream ss = new MemoryStream();
			this.activeStream = cf.CreateEndianStream(ss);
			cf.Root.AddStream((char)0x05 + "SummaryInformation", ss);
			#endregion

			#region Document summary information stream
			MemoryStream dss = new MemoryStream();
			this.activeStream = cf.CreateEndianStream(dss);
			cf.Root.AddStream((char)0x05 + "DocumentSummaryInformation", dss);
			#endregion*/

			cf.Save(stream);
			wbs.Close();
			/*ss.Close();
			dss.Close();*/
		}

		#endregion
	}
}
