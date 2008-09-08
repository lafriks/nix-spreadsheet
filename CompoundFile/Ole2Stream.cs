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
using System.Collections;
using System.IO;

namespace Nix.CompoundFile
{
    /// <summary>
    /// Summary description for Class1.
    /// </summary>
    public class Ole2Stream : Ole2DirectoryEntry
    {
        private Stream stream = null;

        #region Constructor
        public Ole2Stream(string name, Ole2CompoundFile owner, Ole2DirectoryEntry parent) : base(name, EntryType.UserStream, owner, parent)
        {
        }
        #endregion

        #region Stream
		public Stream BaseStream
		{
			get
			{
				return this.stream;
			}
			set
			{
				this.stream = value;
			}
		}
        #endregion

        #region Size
        public override int Size
        {
            get
            {
                return ((int)this.BaseStream.Length);
            }
        }
        #endregion
    }
}
