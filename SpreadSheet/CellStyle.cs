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

// WARNING! This class is generated do not edit!

using System;
using System.Drawing;

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Describes how cell and its data will be displayed.
    /// </summary>
    public class CellStyle : Style
    {
        private Style parent = Style.Default;

        public Style Parent
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

        #region Compare values to parent values
        public bool IsModifiedCellLocked () {
          return (this.CellLocked != this.Parent.CellLocked);
        }

        public bool IsModifiedHiddenFormula () {
          return (this.HiddenFormula != this.Parent.HiddenFormula);
        }

        public bool IsModifiedWrapText () {
          return (this.WrapText != this.Parent.WrapText);
        }

        public bool IsModifiedShrinkToFit () {
          return (this.ShrinkToFit != this.Parent.ShrinkToFit);
        }

        public bool IsModifiedFont () {
          return (this.Font.Equals(this.Parent.Font) == false);
        }

        public bool IsModifiedFormat () {
          return (this.Format != this.Parent.Format);
        }

        public bool IsModifiedIndentLevel () {
          return (this.IndentLevel != this.Parent.IndentLevel);
        }

        public bool IsModifiedTextDirection () {
          return (this.TextDirection != this.Parent.TextDirection);
        }

        public bool IsModifiedWrapTextAtRightBorder () {
          return (this.WrapTextAtRightBorder != this.Parent.WrapTextAtRightBorder);
        }

        public bool IsModifiedHorizontalAlignment () {
          return (this.HorizontalAlignment != this.Parent.HorizontalAlignment);
        }

        public bool IsModifiedVerticalAlignment () {
          return (this.VerticalAlignment != this.Parent.VerticalAlignment);
        }

        public bool IsModifiedJustifyTextAtLastLine () {
          return (this.JustifyTextAtLastLine != this.Parent.JustifyTextAtLastLine);
        }

        public bool IsModifiedRotation () {
          return (this.Rotation != this.Parent.Rotation);
        }

        public bool IsModifiedTopBorderLineStyle () {
          return (this.TopBorderLineStyle != this.Parent.TopBorderLineStyle);
        }

        public bool IsModifiedTopBorderLineColor () {
          return (this.TopBorderLineColor != this.Parent.TopBorderLineColor);
        }

        public bool IsModifiedLeftBorderLineStyle () {
          return (this.LeftBorderLineStyle != this.Parent.LeftBorderLineStyle);
        }

        public bool IsModifiedLeftBorderLineColor () {
          return (this.LeftBorderLineColor != this.Parent.LeftBorderLineColor);
        }

        public bool IsModifiedRightBorderLineStyle () {
          return (this.RightBorderLineStyle != this.Parent.RightBorderLineStyle);
        }

        public bool IsModifiedRightBorderLineColor () {
          return (this.RightBorderLineColor != this.Parent.RightBorderLineColor);
        }

        public bool IsModifiedBottomBorderLineStyle () {
          return (this.BottomBorderLineStyle != this.Parent.BottomBorderLineStyle);
        }

        public bool IsModifiedBottomBorderLineColor () {
          return (this.BottomBorderLineColor != this.Parent.BottomBorderLineColor);
        }

        public bool IsModifiedBackgroundColor () {
          return (this.BackgroundColor != this.Parent.BackgroundColor);
        }

        public bool IsModifiedBackgroundPatternColor () {
          return (this.BackgroundPatternColor != this.Parent.BackgroundPatternColor);
        }

        public bool IsModifiedBackgroundPattern () {
          return (this.BackgroundPattern != this.Parent.BackgroundPattern);
        }
        #endregion

		public override bool Equals(Style other)
		{
			// To be equal parent styles should match
			return base.Equals(other) && other is CellStyle && ((CellStyle)other).Parent.Equals(this.Parent);
		}
    }
}
