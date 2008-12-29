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
using System.Text;

using Nix.CompoundFile;

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Summary description for BIFFRecord.
    /// </summary>
    internal abstract class BIFFRecord
    {
        protected void Write(EndianStream pWriter, int nRecNo, int nRecLen)
        {
            pWriter.Write2(nRecNo);
            pWriter.Write2(nRecLen);
        }

        public abstract void Write(EndianStream pWriter);
 
        public abstract void Read(EndianStream pStream);
    }

    internal abstract class excelValueAttributes : BIFFRecord
    {
        // Row of the value
        private int m_nRow;
        // Column of the value
        private int m_nColumn;
        // All cell attributes held in 3 bytes
        private byte m_nAttr1;
        private byte m_nAttr2;
        private byte m_nAttr3;

        public excelValueAttributes() : this (0, 0)
        {
        }

        public excelValueAttributes(int row, int column)
        {
            this.m_nRow = row;
            this.m_nColumn = column;
            this.m_nAttr1 = (byte)0;
            this.m_nAttr2 = (byte)0;
            this.m_nAttr3 = (byte)0;
        }

        ~excelValueAttributes()
        {
        }

        public void CopyAttributes(excelValueAttributes v)
        {
            this.m_nRow = v.m_nRow;
            this.m_nColumn = v.m_nColumn;
            this.m_nAttr1 = v.m_nAttr1;
            this.m_nAttr2 = v.m_nAttr2;
            this.m_nAttr3 = v.m_nAttr3;
        }

        public void WriteAttributes(EndianStream pWriter)
        {
            pWriter.Write2(this.Row);
            pWriter.Write2(this.Column);
            pWriter.WriteByte(this.m_nAttr1);
            pWriter.WriteByte(this.m_nAttr2);
            pWriter.WriteByte(this.m_nAttr3);
        }

        /* get/set the position of the current item */
        public int Row
        {
            get
            {
                return this.m_nRow;
            }
            set
            {
                this.m_nRow = value;
            }
        }

        public int Column
        {
            set
            {
                this.m_nColumn = value;
            }
            get
            {
                return this.m_nColumn;
            }
        }

        /* Tons and tons of attributes that can be set or not... */
        /*public void setHidden (bool v)
        {
            if (v)
            {
                this.m_nAttr1 |= 0x80;
            }
            else
            {
                this.m_nAttr1 &= unchecked(~0x80);
            }
        }
        public bool getHidden ()
        {
            return (m_nAttr1 & 0x80) != 0;
        }

        public void setLocked (bool v)
        {
            if (v)
            {
                this.m_nAttr1 |= (char)0x40;
            }
            else
            {
                this.m_nAttr1 &= unchecked((char)~0x40);
            }
        }
        public bool getLocked ()
        {
            return (m_nAttr1 & 0x40) != 0;
        }

        public void setShaded (bool v)
        {
            if (v)
            {
                this.m_nAttr3 |= (char)0x80;
            }
            else
            {
                this.m_nAttr3 &= unchecked((char)~0x80);
            }
        }
        public bool getShaded ()
        {
            return (m_nAttr3 & 0x80) != 0;
        }

        public void setBorder (int type)
        {
            // clear existing border
            m_nAttr3 &= unchecked((char)~0x78);
            // set the new border
            m_nAttr3 |= (char)(type & 0x78);
        }
        public int getBorder ()
        {
            return m_nAttr3 & (char)0x78; 
        }

        public void setAlignament (int type)
        {
            // clear previous value
            m_nAttr3 &= unchecked((char)~0x07);
            // set the new value value
            m_nAttr3 |= (char)(type & 0x07);
        }
        public int getAlignament ()
        {
            return m_nAttr3 & (char)0x07;
        }

        public void setFontNum (int v)
        {
            // clear previous value
            m_nAttr2 &= unchecked((char)~0xC0); // was 0xE0
            // set the new value value
            m_nAttr2 |= (char)((v & 0x03) << 5);
        }
        public int getFontNum ()
        {
            return (m_nAttr2 >> 5) & (char)0x03;
        }

        public void setFormatNum (int v)
        {
            // clear previous value
            m_nAttr2 &= unchecked((char)~0x3F);
            // set the new value value
            m_nAttr2 |= (char)(v & 0x3F);
        }
        public int getFormatNum ()
        {
            return m_nAttr2 & (char)0x3F;
        }*/
    }

    public enum SheetType
    {
        WorkBookGlobals = 0x0005,
        VisalBasicModule = 0x0006,
        WorkSheet = 0x0010,
        Chart = 0x0020,
        MacroSheet = 0x0040,
        WorkSpace = 0x0100
    }

    /// <summary>
    /// Write a BOF record
    /// </summary>
    internal class BOFRecord : BIFFRecord
    {
        private SheetType m_nType;
        private int m_nBuildIdentifier;
        private long m_nHistoryFlags;

        public BOFRecord(SheetType nType, int nBuildIdentifier, long nHistoryFlags)
        {
            this.m_nType = nType;
            this.m_nBuildIdentifier = nBuildIdentifier;
            this.m_nHistoryFlags = nHistoryFlags;
        }

        public override void Write(EndianStream pWriter)
        {
            base.Write(pWriter, OPCODE_BOF, 14);
            pWriter.Write2(EXCEL_VERSION);
            pWriter.Write2((int)this.m_nType);
            pWriter.Write2(this.m_nBuildIdentifier);
            pWriter.Write2(DateTime.Today.Year);
            pWriter.Write4(this.m_nHistoryFlags);
            pWriter.Write4(EXCEL_VERSION);
        }
    }

    internal class WINDOW1Record : BIFFRecord
    {
        public WINDOW1Record ()
        {
        }

        public override void Write(EndianStream pWriter)
        {
            base.Write(pWriter, OPCODE_WINDOW1, 18);
            pWriter.Write2(0); // Hpos
            pWriter.Write2(0); // Vpos
            pWriter.Write2(0); // Width
            pWriter.Write2(0); // Height
            pWriter.Write2(0x0008 & 0x0010 & 0x0020); // All visible
            pWriter.Write2(0); // Active worksheet
            pWriter.Write2(0); // First worksheet in tab bar
            pWriter.Write2(1); // Number of selected worksheets
            pWriter.Write2(300); // Width of worksheet tab bar (1/1000 of window width)
        }
    }

    internal class XFRecord : BIFFRecord
    {
        public override void Write(EndianStream pWriter)
        {
            base.Write(pWriter, OPCODE_XF, 12);
            pWriter.WriteByte((byte)0); // font
            pWriter.WriteByte((byte)0); // format
            pWriter.WriteByte((byte)0); // cell type and protection
            pWriter.WriteByte((byte)0xFF); // used attribs
            pWriter.Write2(0); // align, parent
            pWriter.Write2(0); // background
            pWriter.Write4(0xFFFFFFFF); // cell borders
        }
    }

    internal class CODEPAGERecord : BIFFRecord
    {
        private int m_nCodePage = 367; // ASCII

        public CODEPAGERecord (int CodePage)
        {
            this.m_nCodePage = CodePage;
        }

        public override void Write(EndianStream pWriter)
        {
            base.Write(pWriter, OPCODE_CODEPAGE, 2);
            pWriter.Write2(m_nCodePage);
        }
    }

    /// <summary>
    /// Write a number
    /// </summary>
    internal class NUMBERRecord : excelValueAttributes
    {
        private double m_nValue;

        public NUMBERRecord () : this (0)
        {
        }

        public NUMBERRecord (double val)
        {
            this.m_nValue = val;
        }

        public double Value
        {
            set
            {
                this.m_nValue = value;
            }
            get
            {
                return this.m_nValue;
            }
        }

        public override void Write(EndianStream pWriter)
        {
            base.Write(pWriter, OPCODE_NUMBER, 14);
            pWriter.Write2(this.Row);
            pWriter.Write2(this.Column);
            pWriter.Write2(0);
            //base.WriteAttributes(pWriter);
            pWriter.WriteDoubleIEEE(this.m_nValue);
        }
    }

    /// <summary>
    /// Write a label
    /// </summary>
    internal class LABELRecord : excelValueAttributes
    {
        private string m_pchValue;
        private Encoding m_nEncoding;

        public LABELRecord () : this (string.Empty, Encoding.Default)
        {
        }

        public LABELRecord (string val, Encoding encoding)
        {
            this.Value = val;
            this.m_nEncoding = encoding;
        }

        public string Value
        {
            set
            {
                this.m_pchValue = value;
            }
            get
            {
                return this.m_pchValue;
            }
        }

        public override void Write(EndianStream pWriter)
        {
            base.Write(pWriter, OPCODE_LABEL, 8 + this.Value.Length);
            pWriter.Write2(this.Row);
            pWriter.Write2(this.Column);
            pWriter.Write2(0);
            //base.WriteAttributes(pWriter);
            pWriter.Write2(this.Value.Length);
            byte[] bytesSrc = Encoding.Default.GetBytes(this.Value);
            byte[] bytesDest = Encoding.Convert(Encoding.Default, this.m_nEncoding, bytesSrc);
            for (int i = 0; i < bytesDest.Length; i++)
            {
                pWriter.WriteByte(bytesDest[i]);
            }
        }
    }

    /// <summary>
    /// The EOF record
    /// </summary>
    internal class EOFRecord : BIFFRecord
    {
        public override void Write(EndianStream pWriter)
        {
            base.Write(pWriter, OPCODE_EOF, 0);
        }
    }
}
