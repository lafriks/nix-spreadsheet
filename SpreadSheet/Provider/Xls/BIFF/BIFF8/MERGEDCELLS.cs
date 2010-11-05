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

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF8
{
    /// <summary>
    /// This record contains the addresses of merged cell ranges in the current sheet.
    /// </summary>
    internal class MERGEDCELLS : BIFFRecord
	{
		protected override ushort OPCODE {
			get {
                return 0x00E5;
			}
		}

        private List<CellRange> cellRangeList = null;

        public List<CellRange> CellRangeList
        {
            get { return cellRangeList; }
            set { cellRangeList = value; }
        }

        public override void Write(Nix.CompoundFile.EndianStream stream)
        {
            if (this.CellRangeList.Count > 1027)
                throw new IndexOutOfRangeException("CellRangeList can not contain more than 1027 records");
            this.WriteHeader(stream, (ushort)(2 + this.CellRangeList.Count * 8));
            stream.WriteUInt16((ushort)this.CellRangeList.Count);
            foreach (CellRange cr in this.CellRangeList)
            {
                stream.WriteUInt16((ushort)cr.FirstRow);
                stream.WriteUInt16((ushort)cr.LastRow);
                stream.WriteUInt16((ushort)cr.FirstColumn);
                stream.WriteUInt16((ushort)cr.LastColumn);
            }
        }

        public override void Read(Nix.CompoundFile.EndianStream stream)
        {
            throw new NotImplementedException();
        }
    }
}
