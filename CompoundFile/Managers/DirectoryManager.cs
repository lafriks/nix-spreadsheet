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
using System.Collections.Generic;

namespace Nix.CompoundFile.Managers
{
    /// <summary>
    /// Manages directory entries.
    /// </summary>
    internal class DirectoryManager : IEnumerable<Ole2DirectoryEntry>
    {
		private Dictionary<int, Ole2DirectoryEntry> directories = new Dictionary<int, Ole2DirectoryEntry>();
        private int sequence = -1;

		/// <summary>
		/// Register new directory entry.
		/// </summary>
		/// <param name="entry">Directory entry.</param>
		/// <returns>Directory ID.</returns>
        public int Register(Ole2DirectoryEntry entry)
        {
            if (entry == null)
                throw new ArgumentNullException("Directory entry can not be null");
			this.directories.Add(++this.sequence, entry);
			return this.sequence;
        }

		/// <summary>
		/// Directory entry.
		/// </summary>
		/// <value>Directory entry.</value>
        public Ole2DirectoryEntry this[int id]
        {
            get
            {
                if ( ! this.directories.ContainsKey(id))
					throw new ArgumentOutOfRangeException("There is no such directory entry with such ID.");
				return this.directories[id];
            }
        }

		/// <summary>
		/// Gets the directory count.
		/// </summary>
		/// <value>The count.</value>
        public uint Count
        {
			get
			{
				return (uint)this.directories.Count;
			}
        }

		#region IEnumerable<Ole2DirectoryEntry> Members

		/// <summary>
		/// Returns an enumerator that iterates through the collection.
		/// </summary>
		/// <returns>
		/// A <see cref="T:System.Collections.Generic.IEnumerator`1"></see> that can be used to iterate through the collection.
		/// </returns>
		public IEnumerator<Ole2DirectoryEntry> GetEnumerator()
		{
			return this.directories.Values.GetEnumerator();
		}

		#endregion

		#region IEnumerable Members

		/// <summary>
		/// Returns an enumerator that iterates through a collection.
		/// </summary>
		/// <returns>
		/// An <see cref="T:System.Collections.IEnumerator"></see> object that can be used to iterate through the collection.
		/// </returns>
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.directories.Values.GetEnumerator();
		}

		#endregion
	}
}
