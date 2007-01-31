/*
 * Library for writing OLE 2 Compount Document file format.
 * Copyright (C) 2007, Lauris Bukðis-Haberkorns
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

namespace Nix.CompoundFile.Managers
{
	/// <summary>
	/// Manages directory entries.
	/// </summary>
	internal class DirectoryManager
	{
        private ArrayList directories = new ArrayList();
        private bool sync = false;
        private Ole2DirectoryEntry[] dirEntries = null;

        /// <summary>
        /// Register new directory entry.
        /// </summary>
        /// <param name="entry">Directory entry.</param>
        /// <returns>Directory ID.</returns>
        public int Register(Ole2DirectoryEntry entry)
        {
            if (entry == null)
                throw new ArgumentNullException("Directory entry can not be null");
            sync = false;
            return this.directories.Add(entry);
        }

        /// <summary>
        /// List of registered directories.
        /// </summary>
        public Ole2DirectoryEntry[] Directories
        {
            get
            {
                if ( ! this.sync)
                {
                    this.dirEntries = (Ole2DirectoryEntry[])this.directories.ToArray(typeof(Ole2DirectoryEntry));
                    this.sync = true;
                }
                return this.dirEntries;
            }
        }
    }
}
