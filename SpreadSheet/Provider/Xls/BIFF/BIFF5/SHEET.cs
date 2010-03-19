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

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF5
{
	internal class SHEET : BIFFRecord
	{
		/// <summary>
		/// BOF OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x0085;
			}
		}

		public enum SheetType
		{
            Worksheet = 0x0,
			Chart = 0x2,
			VisualBasicModule = 0x6
		}

		[Flags]
		public enum SheetState
		{
			Visible = 0x0,
			Hidden = 0x1,
			VeryHidden = 0x2
		}

        private SheetType type = SheetType.Worksheet;
        private SheetState state = SheetState.Visible;
        private string name = "sheet";
        private uint position = 0;

        /// <summary>
        /// Absolute stream position of the BOF record of the sheet represented by this record
        /// </summary>
        public uint Position
        {
            get { return position; }
            set { position = value; }
        }

        /// <summary>
        /// Sheet name
        /// </summary>
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

		/// <summary>
		/// Sheet type.
		/// </summary>
		public SheetType Type {
			get { return type; }
			set { type = value; }
		}

        /// <summary>
        /// Sheet state.
        /// </summary>
        public SheetState State
        {
            get { return state; }
            set { state = value; }
        }

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write ( EndianStream stream )
		{
			this.WriteHeader(stream, Convert.ToUInt16(6 + BIFFStringHelper.GetStringByteCount(name, false)));
			stream.WriteUInt32(position); // BIFF Version
            stream.WriteByte((byte)this.State); // Sheet state
            stream.WriteByte((byte)this.Type); // Sheet type
            BIFFStringHelper.WriteString(stream, name, false); // Sheet name
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read ( EndianStream stream )
		{
            int lenght = this.ReadHeader(stream);
			Position = stream.Read4(); // BIFF Version
            this.State = (SheetState)stream.ReadByte(); // Sheet state
            this.Type = (SheetType)stream.ReadByte(); // Sheet type
            // TODO: read sheet name
		}
	}
}

