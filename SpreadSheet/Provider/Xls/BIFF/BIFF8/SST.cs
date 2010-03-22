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
using Nix.CompoundFile;
using Nix.SpreadSheet.Provider.Xls.BIFF;

namespace Nix.SpreadSheet.Provider.Xls.BIFF.BIFF8
{
	/// <summary>
	/// SST BIFF record.
	/// </summary>
	internal class SST : BIFFRecord
	{
		private struct SplitItem
		{
			public string String;
			public bool FullString;
			public byte GRBit;
			public ushort StringLength;
			public StringFormating[] StringFormating;
		}

		private class SplitItemCollection
		{
			public List<SplitItem> SplitItemList = new List<SplitItem>();

			public ushort TotalByteCount
			{
				get
				{
					ushort size = 0;
					foreach(SplitItem item in SplitItemList)
						size += BIFFStringHelper.GetStringByteCount(item.String, item.StringFormating, true, item.GRBit, !item.FullString);
					return size;
				}
			}
		}

		public uint TotalCount { get; set; }
		public List<string> StringTable { get; set; }
		private List<SplitItemCollection> RecordBlocks { get; set; }
		
		/// <summary>
		/// SST OPCODE
		/// </summary>
		protected override ushort OPCODE {
			get {
				return 0x00FC;
			}
		}

		public void SplitStringTable ()
		{
			RecordBlocks = new List<SplitItemCollection>();
			RecordBlocks.Add(new SplitItemCollection());
			int left = MaximalRecordLength - 8;
			foreach (string str in StringTable)
			{
				SplitItem item = new SplitItem() { FullString = true, String = str, StringLength = (ushort)str.Length, StringFormating = null, GRBit = BIFFStringHelper.GetGRBIT(str, null, true) };
				ushort len = BIFFStringHelper.GetStringByteCount(item.String, item.StringFormating, true, item.GRBit, !item.FullString);
				while (true)
				{
					// Check if there is space left for string header and at least one character.
					if ((item.FullString && (left < (3 + ((item.GRBit & 0x08) == 0x08 ? 2 : 0) + ((item.GRBit & 0x01) == 0x01 ? 2 : 1))))
						 || (!item.FullString && left < ((item.GRBit & 0x01) == 0x01 ? 2 : 1)))
					{
						// Start new record
						RecordBlocks.Add(new SplitItemCollection());
						left = MaximalRecordLength;
						break;
					}
					if (len <= left)
					{
						RecordBlocks[RecordBlocks.Count - 1].SplitItemList.Add(item);
						left -= len;
						break;
					}
					else
					{
						// Split string into parts
						SplitItem part = new SplitItem() { FullString = false, StringLength = 0, StringFormating = null, GRBit = item.GRBit };
						ushort has_len = 0;
						if (!item.FullString)
							has_len = (ushort)Math.Floor((decimal)left / ((item.GRBit & 0x01) == 0x01 ? 2 : 1));
						else
							has_len = (ushort)Math.Floor((decimal)(left - (3 + ((item.GRBit & 0x08) == 0x08 ? 2 : 0))) / ((item.GRBit & 0x01) == 0x01 ? 2 : 1));
						part.String = item.String.Substring(has_len);
						item.String = item.String.Substring(0, has_len);
						RecordBlocks[RecordBlocks.Count - 1].SplitItemList.Add(item);
						// Start new record
						RecordBlocks.Add(new SplitItemCollection());
						left = MaximalRecordLength;
						item = part;
						len = BIFFStringHelper.GetStringByteCount(item.String, item.StringFormating, true, item.GRBit, !item.FullString);
					}
				}
			}
		}

		public uint GetTotalByteCount()
		{
			if (RecordBlocks == null || RecordBlocks.Count == 0)
				SplitStringTable();
			uint size = 12;
			foreach (SplitItemCollection record in RecordBlocks)
			{
				size += record.TotalByteCount;
			}
			return size + (uint)((RecordBlocks.Count - 1) * 4);
		}

		/// <summary>
		/// Writes BIFF record to the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Write ( EndianStream stream )
		{
			if (RecordBlocks == null || RecordBlocks.Count == 0)
				SplitStringTable();

			this.WriteHeader(stream, (ushort)(RecordBlocks[0].TotalByteCount + 8));
			stream.WriteUInt32(TotalCount); // Total string count
			stream.WriteUInt32((uint)StringTable.Count); // Unique string count

			bool first = true;
			foreach (SplitItemCollection record in RecordBlocks)
			{
				if (!first)
				{
					new CONTINUE() { RecordLength = record.TotalByteCount }.Write(stream);
				}
				foreach (SplitItem item in record.SplitItemList)
				{
					BIFFStringHelper.WriteString(stream, item.String, item.StringFormating, true, item.GRBit, item.StringLength, !item.FullString);
				}
				first = false;
			}
		}

		/// <summary>
		/// Reads BIFF record from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		public override void Read ( EndianStream stream )
		{
			throw new NotImplementedException();
		}
	}
}
