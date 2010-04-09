/*
 * Library for generating spreadsheet documents.
 * Copyright (C) 2008, Lauris Bukðis-Haberkorns <lauris@nix.lv>
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
using System.Security.Cryptography;

namespace Nix.SpreadSheet.Provider.Zip
{
	/// <summary>
	/// Adler-32 Hash algorithm.
	/// </summary>
	internal class Adler32 : HashAlgorithm
	{
        #region Private members and constructor
        private uint adlerA;
        private uint adlerB;

        protected const uint AdlerMod = 65521;

        public Adler32()
        {
            this.Initialize();
        }
        #endregion

        #region Base Adler-32 calculations
        /// <summary>
        /// Initializes an implementation of HashAlgorithm.
        /// </summary>
        public override void Initialize()
        {
            this.adlerA = 1;
            this.adlerB = 0;
        }

        /// <summary>
        /// Core hash computations.
        /// </summary>
        /// <param name="buffer">Data.</param>
        /// <param name="offset">Offset to start at.</param>
        /// <param name="count">Count of bytes to calculate CRC32 for.</param>
        protected override void HashCore(byte[] buffer, int offset, int count)
        {
            // Calculate Adler-32 for given buffer
            uint len = (uint)count;
            while (len > 0)
            {
                uint tlen = (uint)(len > 5550 ? 5550 : len);
                len -= tlen;
                do
                {
                    this.adlerA += buffer[offset++];
                    this.adlerB += this.adlerA;
                } while (--tlen > 0);
                this.adlerA = (this.adlerA & 0xFFFF) + (this.adlerA >> 16) * (65536 - AdlerMod);
                this.adlerB = (this.adlerB & 0xFFFF) + (this.adlerB >> 16) * (65536 - AdlerMod);
            }
        }

        /// <summary>
        /// Calculate final Adler-32 hash value.
        /// </summary>
        /// <returns>Final Adler-32 hash value.</returns>
        protected override byte[] HashFinal()
        {
            uint a = this.adlerA;
            uint b = this.adlerB;
            // It can be shown that a <= 0x1013A here, so a single subtract will do.
            if (a >= AdlerMod)
                a -= AdlerMod;
            // It can be shown that b can reach 0xFFEF1 here.
            b = (b & 0xffff) + (b >> 16) * (65536-AdlerMod);
            if (b >= AdlerMod)
                b -= AdlerMod;
            ulong finalAdler = (b << 16) | a;

            byte [] finalHash = new byte [ 4 ];

            finalHash[0] = (byte) ((finalAdler >> 24) & 0xFF);
            finalHash[1] = (byte) ((finalAdler >> 16) & 0xFF);
            finalHash[2] = (byte) ((finalAdler >>  8) & 0xFF);
            finalHash[3] = (byte) ((finalAdler >>  0) & 0xFF);

            return finalHash;
        }
        #endregion
    }
}
