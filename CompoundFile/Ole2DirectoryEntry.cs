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
using System.Collections;

using Nix.CompoundFile.Managers;

namespace Nix.CompoundFile
{
    internal enum NodeColor
    {
        Red = 0,
        Black = 1
    }
    
    /// <summary>
    /// Directory entry types.
    /// </summary>
    public enum EntryType
    {
        Empty       = 0x00,
        UserStorage = 0x01,
        UserStream  = 0x02,
        LockBytes   = 0x03,
        Property    = 0x04,
        RootStorage = 0x05
    }

	/// <summary>
	/// Summary description for Ole2DirectoryEntry.
	/// </summary>
	public abstract class Ole2DirectoryEntry
	{
	    #region Constructors
        internal Ole2DirectoryEntry(string name, EntryType type, Ole2CompoundFile owner) : this (name, type, owner, null)
        {
        }

        public Ole2DirectoryEntry(string name, EntryType type, Ole2CompoundFile owner, Ole2DirectoryEntry parent)
        {
            this.name = name;
            this.type = type;
            this.parentDir = parent;
            this.owner = owner;
            this.did = owner.DirectoryManager.Register(this);
        }
        #endregion

        #region Entry type
        private EntryType type;

        public EntryType Type
        {
            get
            {
                return this.type;
            }
        }
        #endregion

        #region Owner
        private Ole2CompoundFile owner;

        public Ole2CompoundFile Owner
        {
            get
            {
                return this.owner;
            }
        }
        #endregion

        #region Sector and size
        private int sector = -2;

        internal int Sector
        {
            get
            {
                return this.sector;
            }
            set
            {
                this.sector = value;
            }
        }

        private int size = 0;

        public virtual int Size
        {
            get
            {
                return this.size;
            }
        }

        internal void SetSize(int size)
        {
            this.size = size;
        }
        #endregion

        #region Parent directory (entry)
        private Ole2DirectoryEntry parentDir;

        internal Ole2DirectoryEntry ParentEntry
        {
            get
            {
                return this.parentDir;
            }
        }
        #endregion

	    #region Name
        private string name;

        public string Name
        {
            get
            {
                return this.name;
            }
        }

        public void RenameTo (string name)
        {
            if (this.parentDir == null || this.type == EntryType.RootStorage)
                throw new CompoundFileException("Root storage name can not be changed");
            if (name == string.Empty || name.Length > 31)
                throw new CompoundFileException("Incorect entry name");
            if (this.parentDir != null)
            {
                IEnumerator en = new RedBlackEnumerator(this.parentDir, true);
                while (en.MoveNext())
                {
                    if (((Ole2DirectoryEntry)en.Current).Name == name)
                        throw new CompoundFileException("Entry with such name already exists!");
                }
            }
            this.name = name;
        }
        #endregion

        #region RedBlackNode
        internal Ole2DirectoryEntry(int did)
        {
            this.did = did;
        }

        private NodeColor color = NodeColor.Red;

        internal NodeColor Color
        {
            get
            {
                return this.color;
            }
            set
            {
                this.color = value;
            }
        }

        private int did;

        internal int DID
        {
            get
            {
                return this.did;
            }
            set
            {
                this.did = value;
            }
        }

        private Ole2DirectoryEntry parent = null;
        private Ole2DirectoryEntry left = null;
        private Ole2DirectoryEntry right = null;

        internal Ole2DirectoryEntry Left
        {
            get
            {
                return this.left;
            }
		
            set
            {
                this.left = value;
            }
        }

        internal Ole2DirectoryEntry Right
        {
            get
            {
                return this.right;
            }
		
            set
            {
                this.right = value;
            }
        }

        internal Ole2DirectoryEntry Parent
        {
            get
            {
                return this.parent;
            }
		
            set
            {
                this.parent = value;
            }
        }

        internal virtual Ole2DirectoryEntry Root
        {
            get
            {
                if (this.parent == null)
                    return null;
                Ole2DirectoryEntry r = this.parent;
                while (r.Parent != null)
                {
                    r = r.Parent;
                }
                return r;
            }
        }
        #endregion
	}
}
