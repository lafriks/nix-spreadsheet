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
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Nix.SpreadSheet.Provider.Zip
{
	internal class ZipStream
	{
		private struct Entry
		{
			/// <summary>
			/// Is zip entry compressed or not.
			/// </summary>
			public bool Compress;

			/// <summary>
			/// Zip entry filename.
			/// </summary>
			public string FileName;

			/// <summary>
			/// Last modified date and time.
			/// </summary>
			public DateTime LastModified;

			/// <summary>
			/// Uncompressed file size.
			/// </summary>
			public uint OriginalSize;

			/// <summary>
			/// Compressed file size.
			/// </summary>
			public uint Size;

			/// <summary>
			/// File CRC32 checksum.
			/// </summary>
			public byte[] Crc32;

			/// <summary>
			/// Entry position in stream.
			/// </summary>
			public uint PositionInStream;
			
			/// <summary>
			/// Header size.
			/// </summary>
			public uint HeaderSize;

			/// <summary>
			/// Comment.
			/// </summary>
			public string Comment;
		}

		private Stream stream;

		private List<Entry> entries = new List<Entry>();
		
		/// <summary>
		/// Force strings to be written allways in UTF-8
		/// </summary>
		public bool ForceStringsInUtf8 { get; set; }

		public ZipStream(Stream stream)
		{
			this.ForceStringsInUtf8 = false;
			this.stream = stream;
		}

		#region Header and helper methods
		private bool DetectUtf8(string text)
		{
			foreach (char c in text)
				if (c > 255)
					return true;
			return false;
		}

		private uint DateTimeToDosTime(DateTime dt)
		{
			return (uint)(((dt.Year - 1980) << 25) | (dt.Month << 21) | (dt.Day << 16) | 
							(dt.Hour << 11) | (dt.Minute << 5) | (dt.Second >> 1)); 
		}

		private void WriteFileHeader(ref Entry e)
		{
			e.PositionInStream = (uint)this.stream.Position;

			bool encodeUtf8 = this.ForceStringsInUtf8 || DetectUtf8(e.FileName) /* || DetectUtf8(e.Comment)*/;

			byte[] fileName = (encodeUtf8 ? Encoding.UTF8 : Encoding.GetEncoding(437)).GetBytes(e.FileName);

			// Signature
			this.stream.Write(BitConverter.GetBytes((uint)0x04034b50), 0, 4);
			// Version needed to extract (minimum)
			this.stream.Write(new byte[] {20, 0}, 0, 2);
			// General purpose bit flag
			this.stream.Write(BitConverter.GetBytes((ushort)(encodeUtf8 ? 2048 : 0)), 0, 2);
			// Compression method
			this.stream.Write(BitConverter.GetBytes((ushort)(e.Compress ? 8 : 0)), 0, 2);
			// Last modified date and time
			this.stream.Write(BitConverter.GetBytes(DateTimeToDosTime(e.LastModified)), 0, 4);
			// CRC-32 & size
			if (e.Crc32 != null)
			{
				this.stream.Write(e.Crc32, 0, 4);
				this.stream.Write(BitConverter.GetBytes(e.Size), 0, 4);
				this.stream.Write(BitConverter.GetBytes(e.OriginalSize), 0, 4);
			}
			else
			{
				// Updated later
				this.stream.Write(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }, 0, 12);
			}
			// File name length
			this.stream.Write(BitConverter.GetBytes((ushort)fileName.GetLength(0)), 0, 2);
			// Extra field length
			this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);

			// File name
			this.stream.Write(fileName, 0, fileName.GetLength(0));

			// Set header size
			e.HeaderSize = (uint)(this.stream.Position - e.PositionInStream);
		}
		
		private void UpdateFileHeader(ref Entry e)
		{
            long pos = this.stream.Position;

            this.stream.Position = e.PositionInStream + 14;

            this.stream.Write(e.Crc32, 0, 4);
			this.stream.Write(BitConverter.GetBytes(e.Size), 0, 4);
			this.stream.Write(BitConverter.GetBytes(e.OriginalSize), 0, 4);

            this.stream.Position = pos;
		}

		private void WriteCentralDirectoryHeader(ref Entry e)
		{
			bool encodeUtf8 = this.ForceStringsInUtf8 || DetectUtf8(e.FileName) || DetectUtf8(e.Comment);

			Encoding enc = (encodeUtf8 ? Encoding.UTF8 : Encoding.GetEncoding(437));
			byte[] fileName = enc.GetBytes(e.FileName);
			byte[] comment = enc.GetBytes(e.Comment);
			
			// Central directory file header signature
			this.stream.Write(BitConverter.GetBytes((uint)0x02014b50), 0, 4);
			// Version made by and version needed to extract (minimum)
			this.stream.Write(new byte[] {20, 0, 20, 0}, 0, 4);
			// General purpose bit flag
			this.stream.Write(BitConverter.GetBytes((ushort)(encodeUtf8 ? 2048 : 0)), 0, 2);
			// Compression method
			this.stream.Write(BitConverter.GetBytes((ushort)(e.Compress ? 8 : 0)), 0, 2);
			// Last modified date and time
			this.stream.Write(BitConverter.GetBytes(DateTimeToDosTime(e.LastModified)), 0, 4);
			// File CRC-32
			this.stream.Write(e.Crc32, 0, 4);
			// Compressed size
			this.stream.Write(BitConverter.GetBytes(e.Size), 0, 4);
			// Uncompressed size
			this.stream.Write(BitConverter.GetBytes(e.OriginalSize), 0, 4);
			// File name length
			this.stream.Write(BitConverter.GetBytes((ushort)fileName.GetLength(0)), 0, 2);
			// Extra field length
			this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);
			// File comment length
			this.stream.Write(BitConverter.GetBytes((ushort)comment.GetLength(0)), 0, 2);
			// Disk number where file starts
			this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);
			// Internal file attributes
            this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);
            // External file attributes
            this.stream.Write(BitConverter.GetBytes((uint)0x8100), 0, 4);
            // Relative offset of local file header
            this.stream.Write(BitConverter.GetBytes(e.PositionInStream), 0, 4);
			// File name
			this.stream.Write(fileName, 0, fileName.GetLength(0));
			// File comment
			this.stream.Write(comment, 0, comment.GetLength(0));
		}
		
		private void WriteCentralDirectory()
		{
			uint cdStartPos = (uint)this.stream.Position;

			for (int i = 0; i < this.entries.Count; i++)
			{
				Entry e = this.entries[i];
				this.WriteCentralDirectoryHeader(ref e);
			}

			uint cdSize = (uint)this.stream.Position - cdStartPos;

			// Central directory file header signature
			this.stream.Write(BitConverter.GetBytes((uint)0x06054b50), 0, 4);
			// Number of this disk
			this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);
			// Disk where central directory starts
			this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);
			// Number of central directory records on this disk
            this.stream.Write(BitConverter.GetBytes((ushort)entries.Count), 0, 2);
            // Total number of central directory records
            this.stream.Write(BitConverter.GetBytes((ushort)entries.Count), 0, 2);
            // Size of central directory (bytes)
            this.stream.Write(BitConverter.GetBytes(cdSize), 0, 4);
            // Offset of start of central directory, relative to start of archive
            this.stream.Write(BitConverter.GetBytes(cdStartPos), 0, 4);
            // Comment size
            this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);
            
            this.stream.Flush();
		}
		
		private void WriteStream(ref Entry e, Stream content)
		{
			// Write header
			this.WriteFileHeader(ref e);

			CRC32 crc32 = new CRC32();
			Stream st = ( e.Compress ? new DeflateStream(this.stream, CompressionMode.Compress, true) : this.stream );
            uint fileSize = 0;

            try
			{
	        	byte[] buffer = new byte[8192];
				int r;
				// Write content
				while ( (r = content.Read(buffer, 0, buffer.GetLength(0))) > 0 )
				{
	            	crc32.ComputeHash(buffer, 0, r);
	            	st.Write(buffer, 0, r);
	            	fileSize += (uint)r;
				}
			}
			finally
			{
				if ( e.Compress )
					st.Dispose();
			}
			byte[] hash = (fileSize == 0 ? new byte[] {0, 0, 0, 0} : crc32.Hash);
            Array.Reverse(hash);

            e.OriginalSize = fileSize;
			e.Size = (uint)(this.stream.Position - e.PositionInStream - e.HeaderSize);
			e.Crc32 = hash;
		}
		#endregion

		/// <summary>
		/// Add uncompressed file to zip archive.
		/// </summary>
		/// <param name="fileName">File name.</param>
		/// <param name="content">File content.</param>
		public void AddFileWithStringContent(string fileName, string content)
		{
			AddFileWithStringContent(fileName, content, string.Empty);
		}

		/// <summary>
		/// Add uncompressed file to zip archive.
		/// </summary>
		/// <param name="fileName">File name.</param>
		/// <param name="content">File content.</param>
		/// <param name="comment">Comment.</param>
		public void AddFileWithStringContent(string fileName, string content, string comment)
		{
			byte[] bin = UTF8Encoding.UTF8.GetBytes(content);
            CRC32 crc32 = new CRC32();
            crc32.ComputeHash(bin);
            byte[] hash = crc32.Hash;
            Array.Reverse(hash);
			Entry e = new Entry() { FileName = fileName, Compress = false, OriginalSize = (uint)bin.GetLength(0), Size = (uint)bin.GetLength(0), LastModified = DateTime.Now, Comment = comment, Crc32 = hash };
			// Write header
			this.WriteFileHeader(ref e);
			// Write content
			this.stream.Write(bin, 0, bin.GetLength(0));
			this.stream.Flush();

			this.entries.Add(e);
		}

		/// <summary>
		/// Add and compress stream to zip archive.
		/// </summary>
		/// <param name="fileName">File name.</param>
		/// <param name="content">Stream to read from.</param>
		public void AddStream(string fileName, Stream content)
		{
			AddStream(fileName, content, string.Empty);
		}

		/// <summary>
		/// Add and compress stream to zip archive.
		/// </summary>
		/// <param name="fileName">File name.</param>
		/// <param name="content">Stream to read from.</param>
		/// <param name="comment">Comment.</param>
		public void AddStream(string fileName, Stream content, string comment)
		{
			Entry e = new Entry() { FileName = fileName, Compress = true, LastModified = DateTime.Now, Comment = comment };

			long cpos = content.Position;
			
			this.WriteStream(ref e, content);
			
			if ( e.OriginalSize < e.Size && content.CanSeek )
			{
				e.Compress = false;
				content.Position = cpos;
				this.stream.Position = e.PositionInStream;
				this.stream.SetLength(e.PositionInStream);
				this.WriteStream(ref e, content);
			}

			this.UpdateFileHeader(ref e);

			this.stream.Flush();
			this.entries.Add(e);
		}

		/// <summary>
		/// Write central directory entries and close zip file (Underlaying stream will not be closed).
		/// </summary>
		public void Close()
		{
			this.WriteCentralDirectory();
		}
	}
}
