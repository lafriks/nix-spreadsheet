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

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF8
{
	internal class DIMENSION : BIFFRecord
	{
		/// <summary>
		/// BOF OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x0200;
			}
		}

		

        private uint firstRow = 0;
        private uint lastRow = 0;
        private UInt16 firstCol = 0;
        private UInt16 lastCol = 0;

        /// <summary>
        /// Index to first used row
        /// </summary>
        public uint FirstRow
        {
            get { return firstRow; }
            set { firstRow = value; }
        }
        
        /// <summary>
        /// Index to last used row, increased by 1
        /// </summary>
        public uint LastRow
        {
            get { return lastRow; }
            set { lastRow = value; }
        }

        /// <summary>
        /// Index to first used column
        /// </summary>
        public UInt16 FirstCol
        {
            get { return firstCol; }
            set { firstCol = value; }
        }
        
        /// <summary>
        /// Index to last used column, increased by 1
        /// </summary>
        public UInt16 LastCol
        {
            get { return lastCol; }
            set { lastCol = value; }
        }

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write ( EndianStream stream )
		{
			this.WriteHeader(stream, 14);
			stream.WriteUInt32(firstRow); // first row
			stream.WriteUInt32(lastRow); // last row
			stream.WriteUInt16(firstCol); // first column
			stream.WriteUInt16(lastCol); // last column
			stream.WriteUInt16(0); // Not used
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read ( EndianStream stream )
		{
            // TODO: datu ielasīšana
		}
	}
}

