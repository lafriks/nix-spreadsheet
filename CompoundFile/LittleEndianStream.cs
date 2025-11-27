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
using System.Threading;
using System.Threading.Tasks;
using Mono;

namespace Nix.CompoundFile
{
    /// <summary>
    /// Summary description for LittleEndianWriter.
    /// </summary>
    public class LittleEndianStream : EndianStream
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="stream">Stream to write output to</param>
        public LittleEndianStream(Stream stream) : base(stream)
        {
        }

        /// <summary>
        /// Write 2 bytes ushort type value to the output (little endian order)
        /// </summary>
        /// <param name="value">The value.</param>
        public override void WriteUInt16(ushort value)
        {
            this.Write(DataConverter.LittleEndian.GetBytes(value));
        }

        /// <summary>
        /// Write 2 bytes short type value to the output (little endian order)
        /// </summary>
        /// <param name="value">The value.</param>
        public override void WriteInt16(short value)
        {
            this.Write(DataConverter.LittleEndian.GetBytes(value));
        }

        /// <summary>
        /// Write 4 bytes uint type value to the output (little endian order)
        /// </summary>
        /// <param name="value">The value.</param>
        public override void WriteUInt32(uint value)
        {
            this.Write(DataConverter.LittleEndian.GetBytes(value));
        }

        /// <summary>
        /// Write 4 bytes int type value to the output (little endian order)
        /// </summary>
        /// <param name="value">The value.</param>
        public override void WriteInt32(int value)
        {
            this.Write(DataConverter.LittleEndian.GetBytes(value));
        }

        /// <summary>
        /// Write a 4 byte float type value to the output (little endian order)
        /// </summary>
        /// <param name="value"></param>
        public override void WriteFloatIEEE(float value)
        {
            this.Write(DataConverter.LittleEndian.GetBytes(value));
        }

        /// <summary>
        /// Write a 8 byte double type value to the output (little endian order)
        /// </summary>
        /// <param name=\"value\"></param>
        public override void WriteDoubleIEEE(double value)
        {
            this.Write(DataConverter.LittleEndian.GetBytes(value));
        }

        public override async Task WriteUInt16Async(ushort value, CancellationToken cancellationToken = default)
        {
            await this.WriteAsync(DataConverter.LittleEndian.GetBytes(value), cancellationToken).ConfigureAwait(false);
        }

        public override async Task WriteInt16Async(short value, CancellationToken cancellationToken = default)
        {
            await this.WriteAsync(DataConverter.LittleEndian.GetBytes(value), cancellationToken).ConfigureAwait(false);
        }

        public override async Task WriteUInt32Async(uint value, CancellationToken cancellationToken = default)
        {
            await this.WriteAsync(DataConverter.LittleEndian.GetBytes(value), cancellationToken).ConfigureAwait(false);
        }

        public override async Task WriteInt32Async(int value, CancellationToken cancellationToken = default)
        {
            await this.WriteAsync(DataConverter.LittleEndian.GetBytes(value), cancellationToken).ConfigureAwait(false);
        }

        public override async Task WriteFloatIEEEAsync(float value, CancellationToken cancellationToken = default)
        {
            await this.WriteAsync(DataConverter.LittleEndian.GetBytes(value), cancellationToken).ConfigureAwait(false);
        }

        public override async Task WriteDoubleIEEEAsync(double value, CancellationToken cancellationToken = default)
        {
            await this.WriteAsync(DataConverter.LittleEndian.GetBytes(value), cancellationToken).ConfigureAwait(false);
        }

        public override ushort Read2()
        {
            return DataConverter.LittleEndian.GetUInt16(this.ReadBytes(2), 0);
        }

        public override uint Read4()
        {
            return DataConverter.LittleEndian.GetUInt32(this.ReadBytes(4), 0);
        }

        public override float ReadFloatIEEE()
        {
            return DataConverter.LittleEndian.GetFloat(this.ReadBytes(4), 0);
        }

        public override double ReadDoubleIEEE()
        {
            return DataConverter.LittleEndian.GetDouble(this.ReadBytes(8), 0);
        }
    }
}
