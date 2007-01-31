/*
 * Library for writing OLE 2 Compount Document file format.
 * Copyright (C) 2007, Lauris Bukðis-Haberkorns <lauris@nix.lv>
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
	public abstract class EndianWriter
	{
        protected Stream Stream;

        protected Encoding Encoding;

        #region Constructors
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="stream">Stream to write output to</param>
        public EndianWriter(Stream stream) : this (stream, Encoding.Unicode)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param mame="stream">Stream to write output to</param>
        /// <param name="encoding">String encoding</param>
        public EndianWriter(Stream stream, Encoding encoding)
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
        public abstract void Write2(int value);

        /// <summary>
        /// Write 4 bytes in the output
        /// </summary>
        /// <param name="value"></param>
        public abstract void Write4(long value);

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
    }
}
