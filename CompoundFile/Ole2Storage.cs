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

namespace Nix.CompoundFile
{
	/// <summary>
	/// Summary description for Ole2Storage.
	/// </summary>
	public class Ole2Storage : Ole2DirectoryEntry, IEnumerable
	{
        private RedBlackTree entries;

        #region Constructor
		internal Ole2Storage(Ole2CompoundFile owner) : base("Root Entry", EntryType.RootStorage, owner)
		{
            // Root storage is allways black!
            this.Color = NodeColor.Black;
            this.entries = new RedBlackTree();
        }

		public Ole2Storage(string name, Ole2CompoundFile owner, Ole2DirectoryEntry parent) : base(name, EntryType.UserStorage, owner, parent)
		{
            this.entries = new RedBlackTree();
        }
		#endregion

		#region Indexers
        public Ole2DirectoryEntry this[int DID]
        {
            get
            {
                foreach (Ole2DirectoryEntry entry in this.entries)
                {
                    if (entry.DID == DID)
                        return entry;
                }
                throw new ArgumentOutOfRangeException();
            }
        }

        public Ole2DirectoryEntry this[string name]
        {
            get
            {
                foreach (Ole2DirectoryEntry entry in this.entries)
                {
                    if (entry.Name == name)
                        return entry;
                }
                throw new ArgumentException("Entry with such name does not exist");
            }
        }
        #endregion

        #region Count
        public int Count
        {
            get
            {
                return this.entries.Count();
            }
        }
        #endregion

        #region Add storage and stream methods
        private bool EntryExists(string name)
        {
            foreach (Ole2DirectoryEntry entry in this.entries)
            {
                if (entry.Name == name)
                    return true;
            }
            return false;
        }

        public Ole2Storage AddStorage(string name)
        {
            if (this.EntryExists(name))
                throw new CompoundFileException("Entry with such name already exists!");
            Ole2Storage ne = new Ole2Storage(name, this.Owner, this);
            this.entries.Add(ne);
            return ne;
        }

        public Ole2Stream AddStream(string name, byte[] data)
        {
            if (this.EntryExists(name))
                throw new CompoundFileException("Entry with such name already exists!");
            Ole2Stream ne = new Ole2Stream(name, this.Owner, this);
            ne.SetData(data);
            this.entries.Add(ne);
            return ne;
        }
        #endregion

        #region Remove storage/stream
        private void Remove(Ole2DirectoryEntry obj)
        {
            this.entries.Remove(obj.DID);
        }

        public void Remove(string name)
        {
            this.Remove(this[name]);
        }

        public void RemoveAt(int index)
        {
            this.Remove(this[index]);
        }
        #endregion

        #region Root
        internal override Ole2DirectoryEntry Root
        {
            get
            {
                return this.entries.Root;
            }
        }
        #endregion

        #region IEnumerable Members

        public IEnumerator GetEnumerator()
        {
            return entries.GetEnumerator();
        }

        #endregion

    }
}
