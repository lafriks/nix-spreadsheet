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

namespace Nix.CompoundFile
{
    /// <summary>
    /// Summary description for LittleEndianWriter.
    /// </summary>
    public class LittleEndianWriter : EndianWriter
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="stream">Stream to write output to</param>
        public LittleEndianWriter(Stream stream) : base(stream)
        {
        }

        /// <summary>
        /// Write 2 bytes in the output (little endian order)
        /// </summary>
        /// <param name="v"></param>
        public override void Write2(int value)
        {
            this.WriteByte((byte)((value) & 0xff));
            this.WriteByte((byte)((value >> 8) & 0xff));
        }

        /// <summary>
        /// Write 4 bytes in the output (little endian order)
        /// </summary>
        /// <param name="v"></param>
        public override void Write4(long value)
        {
            this.Write2((int)((value) & 0xffff));
            this.Write2((int)((value >> 16) & 0xffff));
        }

        /// <summary>
        /// Write a 4 byte float in the output
        /// </summary>
        /// <param name="v"></param>
        public override void WriteFloatIEEE(float value)
        {
            byte[] bt = BitConverter.GetBytes(value);
            this.Stream.Write(bt, 0, bt.Length);
        }

        /// <summary>
        /// Write a 8 byte double in the output
        /// </summary>
        /// <param name="v"></param>
        public override void WriteDoubleIEEE(double value)
        {
            byte[] bt = BitConverter.GetBytes(value);
            this.Stream.Write(bt, 0, bt.Length);
        }
    }
}
