﻿/*
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

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Describes how cell and its data will be displayed.
    /// </summary>
    public sealed class Style
    {
        #region Default style
        private static Style defStyle = null;

        public static Style Default
        {
            get
            {
                if (defStyle == null)
                {
                    defStyle = new Style();
                }
                return defStyle;
            }
        }

        internal bool Equals(Style comp)
        {
            return ! (this.wraptext != comp.wraptext || ! this.Font.Equals(comp.Font)
                      || this.shrinkfit != comp.shrinkfit);
        }

        private bool defaultStyle = true;

        public bool IsDefault()
        {
            return this.defaultStyle;
        }
        #endregion

        #region Wrap text
        private bool wraptext = false;

        /// <summary>
        /// Gets or sets if cell value is wrapped.
        /// </summary>
        public bool WrapText
        {
            get
            {
                return this.wraptext;
            }
            set
            {
                this.wraptext = value;
            }
        }
        #endregion
        
        #region Shrink to fit
        private bool shrinkfit = false;

        /// <summary>
        /// Gets or sets if cell value should shrunk to fit the cell.
        /// </summary>
        public bool ShrinkToFit
        {
            get
            {
                return this.shrinkfit;
            }
            set
            {
                this.shrinkfit = value;
            }
        }
        #endregion
        
        #region Font
        private Font font = Font.Default;

        public Font Font
        {
            get
            {
                return this.font;
            }
            set
            {
                this.font = value;
            }
        }
        #endregion
    }
}
