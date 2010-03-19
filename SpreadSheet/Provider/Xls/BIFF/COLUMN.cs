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
using Nix.CompoundFile;

namespace Nix.SpreadSheet.Provider.Xls.BIFF
{
    internal class COLUMN : BIFFRecord
    {
        /// <summary>
        /// EOF record OPCODE.
        /// </summary>
        protected override ushort OPCODE
        {
            get
            {
                return 0x007D;
            }
        }

        private ushort index = 0;

        public ushort Index
        {
            get { return this.index; }
            set { this.index = value; }
        }
        
        private ushort width = 0;

        public ushort Width
        {
            get { return this.width; }
            set { this.width = value; }
        }
        
        private ushort xfIndex = 0;

        public ushort XfIndex
        {
            get { return this.xfIndex; }
            set { this.xfIndex = value; }
        }

        public override void Write(EndianStream stream)
        {
            this.WriteHeader(stream, 12);
            stream.WriteUInt16(Index);
            stream.WriteUInt16(Index);
            stream.WriteUInt16(Width);
            stream.WriteUInt16(XfIndex);
            stream.WriteUInt16(0); // Option flags
            stream.WriteUInt16(0); // unused
        }

        public override void Read(EndianStream stream)
        {
            if (this.ReadHeader(stream) != 0)
                throw new IndexOutOfRangeException("Invalid EOF record");
        }
    }
}

