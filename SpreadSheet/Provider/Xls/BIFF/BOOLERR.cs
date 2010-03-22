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
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
	/// <summary>
	/// BOOLERR record.
	/// </summary>
	internal class BOOLERR : CellBIFFRecord
	{
        /// <summary>
        /// BOOLERR record OPCODE.
        /// </summary>
        protected override ushort OPCODE
        {
            get
            {
                return 0x0205;
            }
        }

        /// <summary>
        /// Value (if boolean 0 - false, 1 - true, or error code)
        /// </summary>
        public byte Value { get; set; }

        /// <summary>
        /// Type (0 - boolean; 1 - error)
        /// </summary>
        public byte Type { get; set; }

        public override void Write(EndianStream stream)
        {
            this.WriteHeader(stream, 8);
            stream.WriteUInt16(RowIndex);
            stream.WriteUInt16(ColIndex);
            stream.WriteUInt16(XfIndex);
            stream.WriteByte(Value);
            stream.WriteByte(Type);
        }

        public override void Read(EndianStream stream)
        {
			throw new NotImplementedException();
        }
	}
}
