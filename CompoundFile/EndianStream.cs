/*
 * Library for writing OLE 2 Compount Document file format.
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
using System.IO;
using System.Text;

namespace Nix.CompoundFile
{
    /// <summary>
    /// Description of BaseWriter.
    /// </summary>
    public abstract class EndianStream
    {
        protected Stream Stream;

        protected Encoding Encoding;

		public long Position
		{
			get
			{
				return Stream.Position;
			}
		}

        #region Constructors
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="stream">Stream to write output to</param>
        public EndianStream(Stream stream) : this (stream, Encoding.Unicode)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param mame="stream">Stream to write output to</param>
        /// <param name="encoding">String encoding</param>
        public EndianStream(Stream stream, Encoding encoding)
        {
            this.Stream = stream;
            this.Encoding = encoding;
        }
        #endregion

        #region Abstract write methods
        /// <summary>
        /// Write 2 bytes in the output
        /// </summary>
        /// <param name="value"></param>
        public abstract void WriteUInt16(ushort value);

        public abstract void WriteInt16(short value);

        /// <summary>
        /// Write 4 bytes in the output
        /// </summary>
        /// <param name="value"></param>
        public abstract void WriteUInt32(uint value);

        public abstract void WriteInt32(int value);

        /// <summary>
        /// Write a 4 byte float in the output
        /// </summary>
        /// <param name="value"></param>
        public abstract void WriteFloatIEEE(float value);

        /// <summary>
        /// Write a 8 byte double in the output
        /// </summary>
        /// <param name="value"></param>
        public abstract void WriteDoubleIEEE(double value);
        #endregion

        #region Abstract read methods
        /// <summary>
        /// Read 2 bytes from the input
        /// </summary>
        public abstract ushort Read2();

        /// <summary>
        /// Read 2 bytes from the input
        /// </summary>
        /// <param name="value"></param>
        public abstract uint Read4();

        /// <summary>
        /// Read a 4 byte float from the input
        /// </summary>
        /// <param name="value"></param>
        public abstract float ReadFloatIEEE();

        /// <summary>
        /// Read a 8 byte double from the input
        /// </summary>
        /// <param name="value"></param>
        public abstract double ReadDoubleIEEE();
        #endregion

        #region Write methods
        public int WriteString(string data)
        {
            byte[] src = Encoding.Default.GetBytes(data);
            byte[] dest = Encoding.Convert(Encoding.Default, this.Encoding, src);
            this.Stream.Write(dest, 0, dest.GetLength(0));
            return dest.GetLength(0);
        }

        public void Write(byte[] buffer, int offset, int count)
        {
            this.Stream.Write(buffer, offset, count);
        }

        public void Write(byte[] value)
        {
        	this.Write(value, 0, value.GetLength(0));
        }

        public void WriteByte(byte value)
        {
            this.Stream.WriteByte(value);
        }

        public void WriteBytes(byte value, int count)
        {
            for(int i = 0; i < count; i++)
                this.Stream.WriteByte(value);
        }
        #endregion

        #region Read methods
        public string ReadString(int byteCount)
        {
        	byte[] src = new byte[byteCount];
        	int sc = this.Stream.Read(src, 0, byteCount);
            byte[] dest = Encoding.Convert(this.Encoding, Encoding.Default, src, 0, sc);
            return this.Encoding.GetString(dest);
        }

        public int Read(byte[] buffer, int offset, int count)
        {
            return this.Stream.Read(buffer, offset, count);
        }
        
        public byte[] ReadBytes(int count)
        {
        	byte[] buffer = new byte[count];
        	if ( this.Stream.Read(buffer, 0, count) == count )
        		return buffer;
        	else
        		throw new EndOfStreamException();
        }

        public byte ReadByte()
        {
            int b = this.Stream.ReadByte();
            if ( b == -1 )
            	throw new EndOfStreamException();
            return (byte)b;
        }
        #endregion
    }
}
