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
	/// Description of MasterSectorAllocationManager.
    /// </summary>
    internal class MasterSectorAllocationManager : IEnumerable<int>
    {
        #region Private variables and constructor
		private List<int> Sectors = new List<int>();
        private SectorAllocationManager SAT;

        public MasterSectorAllocationManager(SectorAllocationManager sat)
        {
            this.SAT = sat;
        }
        #endregion

        #region Allocate SAT sectors
        public int Allocate(uint size)
        {
            return this.Allocate(size, -1);
        }

        public int Allocate(uint size, int val)
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
		/// <summary>
		/// Gets the sector at the specified index.
		/// </summary>
		/// <value></value>
        public int this[int index]
        {
            get
            {
                return this.Sectors[index];
            }
        }

		/// <summary>
		/// Gets the sector count.
		/// </summary>
		/// <value>The count.</value>
        public int Count
        {
			get
			{
				return this.Sectors.Count;
			}
        }
        #endregion

		#region IEnumerable<int> Members

		public IEnumerator<int> GetEnumerator()
		{
			return this.Sectors.GetEnumerator();
		}

		#endregion

		#region IEnumerable Members

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.Sectors.GetEnumerator();
		}

		#endregion
	}
}
