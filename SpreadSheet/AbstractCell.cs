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

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Base class for one or more cells.
    /// </summary>
    public abstract class AbstractCell
    {
        #region Comment
        private Comment comment = null;

        /// <summary>
        /// Gets or sets cell comment.
        /// </summary>
        public Comment Comment
        {
            get
            {
                if (this.comment == null)
                    this.comment = new Comment();
                return this.comment;
            }
            set
            {
                this.comment = value;
            }
        }
        #endregion

        #region Style
        protected CellStyle formatting = null;

        /// <summary>
        /// Returns if cell has set formatting
        /// </summary>
        public bool HasFormatting
        {
            get
            {
                return (this.formatting != null || !this.formatting.Equals(Style.Default));
            }
        }

        /// <summary>
        /// Gets or sets cell style.
        /// </summary>
        public CellStyle Formatting
        {
            get
            {
                if (this.formatting == null)
                    this.formatting = new CellStyle();
                return this.formatting;
            }
            set
            {
                this.formatting = value;
            }
        }
        #endregion

        #region Formula
        private string formula = string.Empty;

        /// <summary>
        /// Formula.
        /// </summary>
        public string Formula
        {
            get
            {
                return this.formula;
            }
            set
            {
                this.formula = value;
            }
        }
        #endregion

        #region Value
        private object val;

        /// <summary>
        /// Gets or sets value on one or more cells. When setting value formula is removed.
        /// </summary>
        public virtual object Value
        {
            get
            {
                if (this.formula != string.Empty)
                    return this.formula;
                else
                    return this.val;
            }
            set
            {
                this.formula = string.Empty;
                this.val = value;
            }
        }

        /// <summary>
        /// Gets display value
        /// </summary>
        public virtual string DisplayValue
        {
            get
            {
                // TODO: Decimal/Date formatting
                if (string.IsNullOrEmpty(this.formula))
                    return Convert.ToString(this.val);
                else
                    // TODO: Evalute formula to get real value
                    return Convert.ToString(this.val);
            }
        }
        #endregion
    }
}
