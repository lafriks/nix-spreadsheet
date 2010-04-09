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
using System.Collections;
using System.IO;
using System.Security.Cryptography;

namespace Nix.SpreadSheet.Provider.Zip
{
	/// <summary>
	/// CRC-32 Hash algorithm.
	/// </summary>
	internal class CRC32 : HashAlgorithm
	{
        #region Static methods and caching
        protected static Hashtable cached;
        protected static bool cache = true;
	
        /// <summary>
        /// Polynomial for CRC-32, big endian.
        /// </summary>
        public const uint DefaultPolynomialBE = 0x04C11DB7;
        /// <summary>
        /// Polynomial for CRC-32, little endian.
        /// </summary>
        public const uint DefaultPolynomialLE = 0xEDB88320;

        /// <summary>
        /// Gets or sets the cache setting.
        /// </summary>
        public static bool Cache
        {
            get
            {
                return cache;
            }
            set
            {
                cache = value;
            }
        }

        /// <summary>
        /// Initialize the cache.
        /// </summary>
        static CRC32()
        {
            cached = Hashtable.Synchronized( new Hashtable() );
        }

        /// <summary>
        /// Clear all cached tables.
        /// </summary>
        public static void ClearCache()
        {
            cached.Clear();
        }

        /// <summary>
        /// Builds a CRC32 table given a polynomial
        /// </summary>
        /// <param name="polynomial">Polynomial</param>
        /// <returns>CRC32 table</returns>
        protected static uint[] BuildTable( uint polynomial )
        {
            uint lcrc;
            uint[] table = new uint[256];

            // 256 values representing ASCII character codes. 
            for (int i = 0; i < 256; i++)
            {
                lcrc = (uint)i;
                for (int j = 0; j < 8; j++)
                {
                    if ((lcrc & 1) != 0)
                        lcrc = (lcrc >> 1) ^ polynomial;
                    else
                        lcrc >>= 1;
                }
                table[i] = lcrc;
            }

            return table;
        }
        #endregion

        #region Private methods and constructors
        protected uint[] table;
        private uint crc;

        /// <summary>
        /// Creates a CRC32 object using the DefaultPolynomial.
        /// </summary>
        public CRC32() : this(CRC32.DefaultPolynomialLE)
        {
        }

        /// <summary>
        /// Creates a CRC32 object using the specified polynomial.
        /// </summary>
        /// <param name="polynomial">Polynomial</param>
        public CRC32(uint polynomial) : this(polynomial, CRC32.Cache)
        {
        }
	
        /// <summary>
        /// Construct a CRC32 object
        /// </summary>
        /// <param name="polynomial">Polynomial</param>
        /// <param name="cache">Specify whether to cache CRC32 table</param>
        public CRC32(uint polynomial, bool cache)
        {
            this.HashSizeValue = 32;

            if ( cached.ContainsKey(polynomial) )
                this.table = (uint [])cached[polynomial];
            else
            {
                this.table = CRC32.BuildTable(polynomial);
                if ( cache )
                    cached.Add(polynomial, this.table);
            }
            this.Initialize();
        }
        #endregion

        #region Base CRC32 calculations
        /// <summary>
        /// Initializes an implementation of HashAlgorithm.
        /// </summary>
        public override void Initialize()
        {
            this.crc = 0xFFFFFFFF;
        }
	
        /// <summary>
        /// Core hash computations.
        /// </summary>
        /// <param name="buffer">Data.</param>
        /// <param name="offset">Offset to start at.</param>
        /// <param name="count">Count of bytes to calculate CRC32 for.</param>
        protected override void HashCore(byte[] buffer, int offset, int count)
        {
            // Calculate CRC32 for given buffer
            for (int i = offset; i < count; i++)
            {
                this.crc = (uint)(this.table[(this.crc & 0xFF) ^ buffer[i]] ^ (this.crc >> 8));
            }
        }
	
        /// <summary>
        /// Calculate final CRC32 hash value.
        /// </summary>
        /// <returns>Final CRC32 hash value.</returns>
        protected override byte[] HashFinal()
        {
            byte [] finalHash = new byte [ 4 ];
            uint finalCRC = ~ this.crc;
		
            finalHash[0] = (byte) ((finalCRC >> 24) & 0xFF);
            finalHash[1] = (byte) ((finalCRC >> 16) & 0xFF);
            finalHash[2] = (byte) ((finalCRC >>  8) & 0xFF);
            finalHash[3] = (byte) ((finalCRC >>  0) & 0xFF);
		
            return finalHash;
        }
        #endregion
    }
}
