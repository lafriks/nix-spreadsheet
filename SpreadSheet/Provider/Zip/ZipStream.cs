using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

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
			public uint? Crc32;

			/// <summary>
			/// Header length.
			/// </summary>
			public uint HeaderLength;

			/// <summary>
			/// Comment.
			/// </summary>
			public string Comment;
		}

		private Stream stream;

		private List<Entry> entries = new List<Entry>();

		public ZipStream(Stream stream)
		{
			this.stream = stream;
		}

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

		private void WriteFileHeader(Entry e)
		{
			long hStartPos = this.stream.Position;

			bool encodeUtf8 = DetectUtf8(e.FileName)/* || DetectUtf8(e.Comment)*/;

			byte[] fileName = (encodeUtf8 ? Encoding.UTF8 : Encoding.ASCII).GetBytes(e.FileName);

			// Signature
			this.stream.Write(BitConverter.GetBytes(0x04034b50), 0, 4);
			// Version needed to extract (minimum)
			this.stream.Write(new byte[] {20, 0}, 0, 2);
			// Encode in utf8 or ascii
			this.stream.Write(BitConverter.GetBytes((ushort)(encodeUtf8 ? 2048 : 0)), 0, 2);
			// Compression method
			this.stream.Write(BitConverter.GetBytes((ushort)(e.Compress ? 8 : 0)), 0, 2);
			// Last modified date and time
			this.stream.Write(BitConverter.GetBytes(DateTimeToDosTime(e.LastModified)), 0, 4);
			// CRC & size
			if (e.Crc32.HasValue)
			{
				this.stream.Write(BitConverter.GetBytes(e.Crc32.Value), 0, 4);
				this.stream.Write(BitConverter.GetBytes(e.OriginalSize), 0, 4);
				this.stream.Write(BitConverter.GetBytes(e.Size), 0, 4);
			}
			else
			{
				// Updated later
				this.stream.Write(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }, 0, 12);
			}
			// Filename length
			this.stream.Write(BitConverter.GetBytes((ushort)fileName.GetLength(0)), 0, 2);
			// Extra field length
			this.stream.Write(BitConverter.GetBytes((ushort)0), 0, 2);

			// Filename
			this.stream.Write(fileName, 0, fileName.GetLength(0));

			// Set header lenght
			e.HeaderLength = (uint)(this.stream.Position - hStartPos);
		}

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
			Entry e = new Entry() { FileName = fileName, Compress = false, OriginalSize = (uint)bin.GetLength(0), Size = (uint)bin.GetLength(0), LastModified = DateTime.Now, Comment = comment };
			entries.Add(e);
			// Write header
			WriteFileHeader(e);
			// Write content
			long eStartPos = stream.Position;
			stream.Write(bin, 0, bin.GetLength(0));
		}
	}
}
