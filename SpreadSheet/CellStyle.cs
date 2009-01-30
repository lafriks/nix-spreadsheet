/*
 * Library for generating spreadsheet documents.
 * Copyright (C) 2008, Lauris Bukšis-Haberkorns <lauris@nix.lv>
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
using System.Drawing;

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Describes how cell and its data will be displayed.
    /// </summary>
    public class CellStyle : ExtendedStyle
    {
    	#region Parent style
		private ExtendedStyle style = null; //ExtendedStyle.Default;

		public ExtendedStyle Style
		{
			get { return this.style; }
			set { this.style = value; }
		}
    	#endregion

        #region Default style
        private static CellStyle defStyle = null;

        public static CellStyle Default
        {
            get
            {
                if (defStyle == null)
                {
                    defStyle = new CellStyle();
                }
                return defStyle;
            }
        }

        internal bool Equals(CellStyle comp)
        {
            return ! (this.WrapText != comp.WrapText || ! this.Font.Equals(comp.Font)
                      || this.ShrinkToFit != comp.ShrinkToFit);
        }

        private bool defaultStyle = true;

        public bool IsDefault()
        {
            return this.defaultStyle;
        }
        #endregion
    }
}
