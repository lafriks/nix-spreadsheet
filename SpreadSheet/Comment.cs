/*
 * Library for creating Microsoft Excel files.
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

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Comment.
    /// </summary>
    public class Comment
    {
        private string text = string.Empty;

        /// <summary>
        /// Gets or sets comment text assigned to cell.
        /// </summary>
        public string Text
        {
            get
            {
				return this.text;
            }
            set
            {
				this.text = value;
            }
        }

		private bool visible = false;

		/// <summary>
        /// Gets or sets value indicating whether comment is visible.
        /// </summary>
        public bool Visible
        {
            get
            {
                return this.visible;
            }
            set
            {
                this.visible = value;
            }
        }
    }
}
