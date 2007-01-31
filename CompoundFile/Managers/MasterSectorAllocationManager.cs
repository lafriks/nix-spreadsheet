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
	/// Description of ShortSectorAllocationManager.
	/// </summary>
	internal class MasterSectorAllocationManager
	{
	    #region Private variables and constructor
	    private ArrayList Sectors = new ArrayList();
        private SectorAllocationManager SAT;

        private bool sync = false;
        private int[] allocations;

	    public MasterSectorAllocationManager(SectorAllocationManager sat)
		{
            this.SAT = sat;
		}
	    #endregion

	    #region Allocate SAT sectors
        public int Allocate(int size)
        {
            return this.Allocate(size, -1);
        }

        public int Allocate(int size, int val)
        {
            this.SAT.Allocation += new SectorEventHandler(SAT_Allocation);
            int res = this.SAT.Allocate(size, val);
            this.SAT.Allocation -= new SectorEventHandler(SAT_Allocation);
            return res;
        }

        private void SAT_Allocation(object sender, int index)
        {
            this.Sectors.Add(index);
        }
        #endregion

        #region Master table allocations
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
    }
}
