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
	/// <summary>
	/// Microsoft Excel binary file format provider.
	/// </summary>
	public class XlsFileFormatProvider : IFileFormatProvider
	{
		private EndianStream activeStream = null;
		private void Write(BIFFRecord record)
		{
			if ( activeStream == null )
				throw new ArgumentNullException("activeStream");
			record.Write(activeStream);
		}

        private ColorPalette palette = new ColorPalette();        

		private List<Font> fontTable = new List<Font>();

		private Dictionary<string, uint> stringTable = new Dictionary<string, uint>();

		/// <summary>
		/// Find font index.
		/// </summary>
		/// <param name="font">Font object.</param>
		/// <returns>Index in font table.</returns>
		protected int FindFontIndex(Font font)
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
		
		/// <summary>
		/// Build used font table.
		/// </summary>
		/// <param name="document">Spreadsheet document to search used fonts in.</param>
		protected void BuildFontTable ( SpreadSheetDocument document )
		{
			// first default font for setting column widths
            fontTable.Add(Font.Default);

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

		/// <summary>
		/// Find format index.
		/// </summary>
		/// <param name="format">Format.</param>
		/// <returns>Format index in format table.</returns>
		protected ushort FindFormatIndex(string format)
		{
			foreach ( ushort idx in formatTable.Keys )
				if ( formatTable[idx] == format )
					return idx;
			throw new IndexOutOfRangeException("Format not found");
		}

		private ushort formatTableSeq = 0;

		/// <summary>
		/// Build used format table.
		/// </summary>
		/// <param name="document">Spreadsheet document to search used formats in.</param>
		protected void BuildFormatTable ( SpreadSheetDocument document )
		{
			// Add default formats
			this.formatTable.Add(0, "General");
			this.formatTable.Add(1, "0");
			this.formatTable.Add(2, "0.00");
			this.formatTable.Add(3, "#,##0");
			this.formatTable.Add(4, "#,##0.00");
			if ( document.Locale.TwoLetterISORegionName == "LV")
			{
				this.formatTable.Add(5, "#,##0\\ \"Ls\";\\-#,##0\\ \"Ls\"");
				this.formatsToWrite.Add(5);
				this.formatTable.Add(6, "#,##0\\ \"Ls\";[Red]\\-#,##0\\ \"Ls\"");
				this.formatsToWrite.Add(6);
				this.formatTable.Add(7, "#,##0.00\\ \"Ls\";\\-#,##0.00\\ \"Ls\"");
				this.formatsToWrite.Add(7);
				this.formatTable.Add(8, "#,##0.00\\ \"Ls\";[Red]\\-#,##0.00\\ \"Ls\"");
				this.formatsToWrite.Add(8);
			}
			else
			{
				this.formatTable.Add(5, "($#,##0_);($#,##0)");
				this.formatTable.Add(6, "($#,##0_);[Red]($#,##0)");
				this.formatTable.Add(7, "($#,##0.00_);($#,##0.00)");
				this.formatTable.Add(8, "($#,##0.00_);[Red]($#,##0.00)");
			}
			this.formatTable.Add(9, "0%");
			this.formatTable.Add(10, "0.00%");
			this.formatTable.Add(11, "0.00E+00");
			this.formatTable.Add(12, "# ?/?");
			this.formatTable.Add(13, "# ??/??");
			this.formatTable.Add(14, "m/d/yy");
			this.formatTable.Add(15, "d-mmm-yy");
			this.formatTable.Add(16, "d-mmm");
			this.formatTable.Add(17, "mmm-yy");
			this.formatTable.Add(18, "h:mm AM/PM");
			this.formatTable.Add(19, "h:mm:ss AM/PM");
			this.formatTable.Add(20, "h:mm");
			this.formatTable.Add(21, "h:mm:ss");
			this.formatTable.Add(22, "m/d/yy h:mm");
			this.formatTable.Add(37, "(#,##0_);(#,##0)");
			this.formatTable.Add(38, "(#,##0_);[Red](#,##0)");
			this.formatTable.Add(39, "(#,##0.00_);(#,##0.00)");
			this.formatTable.Add(40, "(#,##0.00_);[Red](#,##0.00)");
			if ( document.Locale.TwoLetterISORegionName == "LV")
			{
				this.formatTable.Add(41, "_-* #,##0\\ _L_s_-;\\-* #,##0\\ _L_s_-;_-* \"-\"\\ _L_s_-;_-@_-");
				this.formatsToWrite.Add(41);
				this.formatTable.Add(42, "_-* #,##0\\ \"Ls\"_-;\\-* #,##0\\ \"Ls\"_-;_-* \"-\"\\ \"Ls\"_-;_-@_-");
				this.formatsToWrite.Add(42);
				this.formatTable.Add(43, "_-* #,##0.00\\ _L_s_-;\\-* #,##0.00\\ _L_s_-;_-* \"-\"??\\ _L_s_-;_-@_-");
				this.formatsToWrite.Add(43);
				this.formatTable.Add(44, "_-* #,##0.00\\ \"Ls\"_-;\\-* #,##0.00\\ \"Ls\"_-;_-* \"-\"??\\ \"Ls\"_-;_-@_-");
				this.formatsToWrite.Add(44);
			}
			else
			{
				this.formatTable.Add(41, "_(* #,##0_);_(* (#,##0);_(* \" - \"_);_(@_)");
				this.formatTable.Add(42, "_($* #,##0_);_($* (#,##0);_($* \" - \"_);_(@_)");
				this.formatTable.Add(43, "_(* #,##0.00_);_(* (#,##0.00);_(* \" - \"??_);_(@_)");
				this.formatTable.Add(44, "_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)");
			}
			this.formatTable.Add(45, "mm:ss");
			this.formatTable.Add(46, "[h]:mm:ss");
			this.formatTable.Add(47, "mm:ss.0");
			this.formatTable.Add(48, "##0.0E+0");
			this.formatTable.Add(49, "@");

			this.formatTableSeq = 164;
			
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

		/// <summary>
		/// Find style index.
		/// </summary>
		/// <param name="style">Style.</param>
		/// <returns>Style index in format table.</returns>
		public int FindStyleIndex(Style style)
		{
			for (int i = 0; i < styleTable.Count; i++)
			{
				if ( styleTable[i].Equals(style) )
					return i + 16;
			}
			return -1;
		}

		/// <summary>
		/// Build style table.
		/// </summary>
		/// <param name="document">Spreadsheet document to search used styles in.</param>
		protected void BuildStyleTable(SpreadSheetDocument document)
		{
            this.styleTable.Add(new Style());
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
            if (this.styleTable.Count == 0)
            {
                this.styleTable.Add(Style.Default);
            }
		}

		/// <summary>
		/// Get value type.
		/// </summary>
		/// <param name="cell">Cell.</param>
		/// <returns>-1 - empty; 0 - string; 1 - number; 2 - boolean or error.</returns>
		protected int GetValueType(Cell cell)
		{
			if (cell.Value == null || (cell.Value is string && (string)cell.Value == string.Empty))
				return -1;
			if (cell.Value is ErrorCode || cell.Value is bool)
				return 2;
			if (cell.Value is int || cell.Value is float || cell.Value is double ||
                	cell.Value is decimal || cell.Value is long || cell.Value is short)
            	return 1;
			return 0;
		}

		/// <summary>
		/// Format cell value.
		/// </summary>
		/// <param name="cell">Cell.</param>
		/// <returns>Formated cell value.</returns>
		protected string FormatValue(Cell cell)
		{
			// TODO: Better value convertation to string
			return Convert.ToString(cell.Value);
		}
		
		protected byte GetErrorCodeValue(ErrorCode code)
		{
			switch (code)
			{
				default:
				case ErrorCode.Null:
					return 0x00;
				case ErrorCode.DivisionByZero:
					return 0x07;
				case ErrorCode.WrongType:
					return 0x0F;
				case ErrorCode.IllegalCellReference:
					return 0x17;
				case ErrorCode.WrongFunctionOrRange:
					return 0x1D;
				case ErrorCode.ValueRangeOverflow:
					return 0x24;
				case ErrorCode.ArgumentOrFunctionNotAvailable:
					return 0x2A;
			}
		}

		/// <summary>
		/// Build strings table.
		/// </summary>
		/// <param name="document">Spreadsheet document to search strings in.</param>
		/// <returns>Total count of strings in document.</returns>
		protected uint BuildStringTable(SpreadSheetDocument document)
		{
			uint count = 0;
			// Build string table
			foreach ( Sheet sheet in document )
			{
				foreach ( Row row in sheet )
				{
					foreach ( Cell cell in row )
					{
						if ( this.GetValueType(cell) == 0 )
						{
							count++;
							string val = this.FormatValue(cell);
							if ( ! stringTable.ContainsKey(val) )
								stringTable.Add(val, (uint)stringTable.Count);
						}
					}
				}
			}
			return count;
		}
		
		/// <summary>
		/// Find string in table.
		/// </summary>
		/// <param name="value">String.</param>
		/// <returns>Index of string in table.</returns>
		protected uint FindStringTableIndex(string value)
		{
			if ( ! stringTable.ContainsKey(value) )
				throw new ArgumentOutOfRangeException("value");
			return stringTable[value];
		}

		#region IFileFormatProvider Members
		private void WriteRowBlock(List<Row> rows)
		{
			foreach(Row row in rows)
			{
				this.Write(new ROW() { Row = row });
			}
			// Write cells
			foreach (Row row in rows)
			{
				foreach (Cell cell in row)
				{
					switch (this.GetValueType(cell))
					{
						case 1:
							this.Write(new NUMBER()
							{
								ColIndex = (ushort)cell.ColumnIndex,
								RowIndex = (ushort)cell.RowIndex,
								XfIndex = (ushort)FindStyleIndex(cell.Formatting),
								Value = Convert.ToDouble(cell.Value)
							});
							break;
						case 2:
							this.Write(new BOOLERR()
							{
								ColIndex = (ushort)cell.ColumnIndex,
								RowIndex = (ushort)cell.RowIndex,
								XfIndex = (ushort)FindStyleIndex(cell.Formatting),
								Value = (byte)(cell.Value is bool ? ((bool)cell.Value ? 1 : 0) : GetErrorCodeValue((ErrorCode)cell.Value)),
								Type = (byte)(cell.Value is bool ? 0 : 1)
							});
							break;
						case -1:
							this.Write(new BLANK()
							{
								ColIndex = (ushort)cell.ColumnIndex,
								RowIndex = (ushort)cell.RowIndex,
								XfIndex = (ushort)FindStyleIndex(cell.Formatting),
							});
							break;
						default:
							this.Write(new LABELSST()
							{
								ColIndex = (ushort)cell.ColumnIndex,
								RowIndex = (ushort)cell.RowIndex,
								XfIndex = (ushort)FindStyleIndex(cell.Formatting),
								IndexToSST = FindStringTableIndex(FormatValue(cell))
							});
							break;
					}
				}
			}
			// TODO: Write DBCELL
		}

		/// <summary>
		/// Save document to stream.
		/// </summary>
		/// <param name="document">Spreadsheet document.</param>
		/// <param name="stream">Stream to write to.</param>
		public void Save ( SpreadSheetDocument document, System.IO.Stream stream )
		{
			Ole2CompoundFile cf = new Ole2CompoundFile();

			this.BuildFontTable(document);
			this.BuildFormatTable(document);
			this.BuildStyleTable(document);
			uint totalStringCount = this.BuildStringTable(document);
			
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
                this.Write(new FONT() { Font = font, Palette = palette });
			// Format table
			foreach ( ushort fi in this.formatsToWrite )
				this.Write(new FORMAT() { Index = fi, Format = this.formatTable[fi] });
            // Default XFs
            for (int j = 0; j < 15; j++)
            {
                XF r = new XF(){Style = this.styleTable[0],
                                FontIndex = (ushort)FindFontIndex(this.styleTable[0].Font),
                                FormatIndex = FindFormatIndex(this.styleTable[0].Format),
                                Palette = palette};
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
				XF r = new XF() { Style = s, FontIndex = (ushort)FindFontIndex(s.Font), FormatIndex = FindFormatIndex(s.Format), Palette = palette };
				if ( s is CellStyle )
                    r.ParentStyleIndex = (ushort?)FindStyleIndex(((CellStyle)s).Parent);
				this.Write(r);
			}
			
			this.Write(new PALETTE() {Palette = palette});

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

                this.Write(new PAGESETUP() { PageSettings = sheet.PageSettings });

				foreach (Column col in sheet.Columns)
				{
					// Convert pixels to MS mystical units
					ushort width = Convert.ToUInt16(Math.Round((double)col.Width/7 * 256));
					this.Write(new COLUMN() { Index = (ushort)col.ColumnIndex, Width = (ushort)width, XfIndex = (ushort)FindStyleIndex(col.Formatting) });
				}
				
				this.Write(new DIMENSION() { FirstCol = (ushort)sheet.FirstColumn, FirstRow = (uint)sheet.FirstRow,
				           					 LastCol = (ushort)(sheet.LastColumn + 1), LastRow = (uint)(sheet.LastRow + 1) });

				// Write rows in 32 row blocks
				List<Row> rows = new List<Row>();
				foreach ( Row row in sheet )
				{
					if ( rows.Count > 0 && rows[rows.Count - 1].RowIndex != row.RowIndex - 1 )
					{
						WriteRowBlock(rows);
						rows.Clear();
					}
					rows.Add(row);
					if ( rows.Count == 32 )
					{
						WriteRowBlock(rows);
						rows.Clear();
					}
				}
				if ( rows.Count > 0 )
				{
					WriteRowBlock(rows);
					rows.Clear();
				}
				this.Write(new WINDOW2());

                int mergedCellRangeCount = sheet.mergedCells.Count;
                int mi = 0;

                // Write merged cell ranges
                while (mergedCellRangeCount > 0)
                {
                    this.Write(new MERGEDCELLS() { CellRangeList = sheet.mergedCells.GetRange(mi, Math.Min(1027, mergedCellRangeCount)) });

                    mi += 1027;
                    mergedCellRangeCount -= 1027;
                }

				this.Write(new EOF());
			}

            uint settingsLength = (uint)(document.SheetCount * 10 // Sheet base data length (without sheet name)
                                         + sheetNamesLength // Sheet name length in unicode
                                         + 4); // EOF
			
			// Calculate SST length
			SST sst = null;
			if ( stringTable.Count > 0 )
				sst = new SST() { TotalCount = totalStringCount, StringTable = stringTable };

            this.activeStream = mainStream;
			uint currentPosition = (uint)wbs.Length + settingsLength + (sst != null ? sst.GetTotalByteCount() : 0);
            int i = 0;
            // Sheets' headers
            foreach(Sheet sheet in document)
            {
                this.Write(new SHEET() {Name = sheet.Name, Position = currentPosition});
                currentPosition += (uint)sheetStreams[i].Length;
                i++;
            }
			// Write Shared String Table
			if ( sst != null )
            	this.Write(sst);
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
