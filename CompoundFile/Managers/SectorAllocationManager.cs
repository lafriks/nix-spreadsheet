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
using System.Collections;

namespace Nix.CompoundFile.Managers
{
    internal delegate void SectorEventHandler(object sender, int index);

    /// <summary>
    /// Description of SectorAllocationManager.
    /// </summary>
    internal class SectorAllocationManager
    {
        #region Private variables
        private int SectorSize = 0;
        private ArrayList Sectors = new ArrayList();
        private ArrayList SectorsData = new ArrayList();

        private bool sync = false;
        private int[] allocations;
        #endregion

        #region Events
        public event SectorEventHandler Allocation;

        public void OnAllocation (object sender, int index)
        {
            if (this.Allocation != null)
            {
                this.Allocation(sender, index);
            }
        }
        #endregion

        #region Constructor
        public SectorAllocationManager(int size)
        {
            this.SectorSize = size;
        }
        #endregion

        #region Allocate methds
        private int FindFirstFree()
        {
            return this.FindFirstFree(0);
        }

        private int FindFirstFree(int from)
        {
            for (int i = from; i < this.Sectors.Count; i++)
            {
                if ((int)this.Sectors[i] == -1)
                    return i;
            }
            return -1;
        }

        public int Allocate(int size)
        {
            return this.Allocate(size, -1);
        }

        public int Allocate(int size, int val)
        {
            this.sync = false;
            int f = this.FindFirstFree();
            int n = f;
            int nx = -2;
            int count = Convert.ToInt32(Math.Ceiling((double)size / this.SectorSize));
            for (int i = 0; i < count; i++)
            {
                nx = this.FindFirstFree(n + 1);
                if (n == -1)
                {
                    n = this.Sectors.Add(-1);
                    // keep in sync
                    this.SectorsData.Add(null);
                }
                if (f == -1)
                    f = n;
                this.Sectors[n] = (val == -1 ? (nx == -1 ? -2 : nx) : val);
                this.OnAllocation(this, n);
            }
            if (nx >= 0)
            {
                this.Sectors[nx] = (val == -1 ? -2 : val);
                this.OnAllocation(this, nx);
            }
            return f;
        }
        #endregion

        #region Allocations
        public int[] Allocations
        {
            get
            {
                if ( ! this.sync)
                {
                    this.allocations = (int[])this.Sectors.ToArray(typeof(int));
                    this.sync = true;
                }
                return this.allocations;
            }
        }
        #endregion

        #region Allocate stream data
        public void AllocateStream(int start, byte[] data)
        {
            this.AllocateStream(start, data, 0);
        }

        public void AllocateStream(int start, byte[] data, byte def)
        {
            int offset = 0;
            int steamSize = data.GetLength(0);
            while (start > -1)
            {
                byte[] nd = new byte[this.SectorSize];
                // Copy one sector data
                for (int i = 0; i < this.SectorSize; i++)
                    if (i + offset < steamSize)
                        nd[i] = data[i + offset];
                    else
                        nd[i] = def;
                this.SectorsData[start] = nd;
                offset += this.SectorSize;
                // Go to next sector
                start = (int)this.Sectors[start];
            }
        }
        #endregion

        #region Write sectors to stream
        public void WriteData(Stream writer)
        {
            foreach(byte[]sector in this.SectorsData)
            {
                writer.Write(sector, 0, this.SectorSize);
            }
        }
        #endregion
    }
}
