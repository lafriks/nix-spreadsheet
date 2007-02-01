/*
 * Library for creating Microsoft Excel files.
 * Copyright (C) 2007, Lauris Bukšis-Haberkorns <lauris@nix.lv>
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
using System.Collections;
using System.IO;
using System.Text;

using Nix.CompoundFile;

namespace Nix.SpreadSheet
{
    public enum CellType
    {
        None = 0,
        Number = 1,
        String = 2
    }

    /// <summary>
    /// Summary description for Excel.
    /// </summary>
    public class SpreadSheet
    {
        public const int MaxRows = 65536;
        public const int MaxColumns = 256;

        private Hashtable m_table = new Hashtable();
	    
        public SpreadSheet()
        {
        }

        public void Save(string fileName)
        {
            StreamWriter sw = new StreamWriter(fileName);
            this.Save(sw.BaseStream);
            sw.Close();
        }

        public void Save(Stream pWriter)
        {
            Ole2CompoundFile file = new Ole2CompoundFile();

            #region Workbook
            MemoryStream mem = new MemoryStream();

            EndianWriter writer = new LittleEndianWriter(mem);

            BOFRecord biffBOF = new BOFRecord(SheetType.WorkBookGlobals, 0, 0);
            // It's allways UTF-16
            CODEPAGERecord biffCODEPAGE = new CODEPAGERecord(1200);
            WINDOW1Record biffWINDOW1 = new WINDOW1Record();
            EOFRecord biffEOF = new EOFRecord();

            biffBOF.Write(writer);
            biffCODEPAGE.Write(writer);
            biffWINDOW1.Write(writer);


            biffEOF.Write(writer);

            file.Root.AddStream("Book", mem.ToArray());
            file.Root.AddStream("Workbook", mem.ToArray());
            #endregion

            /*BOFRecord biffBOF = new BOFRecord(BIFFRecord.EXCEL_VERSION, BIFFRecord.TYPE_WORKSHEET);
            CODEPAGERecord biffCODEPAGE = new CODEPAGERecord(this.m_encoding.CodePage);
            EOFRecord biffEOF = new EOFRecord();

            biffBOF.Write(writer);
            biffCODEPAGE.Write(writer);

            XFRecord x = new XFRecord();
            x.Write(writer);

            foreach (string key in this.m_table.Keys)
            {
                ExcelCell cell = (ExcelCell)this.m_table[key];
                if (cell.getType() == CellType.Number)
                {
                    cell.Write(writer);
                }
            }

            foreach (string key in this.m_table.Keys)
            {
                ExcelCell cell = (ExcelCell)this.m_table[key];
                if (cell.getType() == CellType.String)
                {
                    cell.Write(writer);
                }
            }

            biffEOF.Write(writer);*/

            file.Save(pWriter);
        }

        public Cell this[int row, int column]
        {
            get
            {
                if (row >= MaxRows || column >= MaxColumns)
                    throw new ArgumentOutOfRangeException();
                string key = Utils.CellName(row, column);
                if (this.m_table.ContainsKey(key))
                {
                    return (Cell)this.m_table[key];
                }
                else
                {
                    Cell nc = new Cell(row, column);
                    this.m_table.Add(key, nc);
                    return nc;
                }
            }
        }

        public Cell this[string name]
        {
            get
            {
                if (this.m_table.ContainsKey(name))
                {
                    return (Cell)this.m_table[name];
                }
                else
                {
                    Cell nc = new Cell(name);
                    if (nc.RowIndex >= MaxRows || nc.ColumnIndex >= MaxColumns)
                        throw new ArgumentOutOfRangeException();
                    this.m_table.Add(nc.Name, nc);
                    return nc;
                }
            }
        }
    }
}
