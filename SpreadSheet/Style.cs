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
    public sealed class Style
    {
    	#region Cell protection
		private bool cellLocked = false;
		
		public bool CellLocked {
			get { return cellLocked; }
			set { cellLocked = value; }
		}

		private bool hiddenFormula = false;
		
		public bool HiddenFormula {
			get { return hiddenFormula; }
			set { hiddenFormula = value; }
		}
		#endregion

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
        private Font font = new Font();

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

        #region Format
        private string format = string.Empty;

		public string Format {
			get { return format; }
			set { format = value; }
		}
        #endregion

        #region Alignment
		private bool wrapTextAtRightBorder = true;

		public bool WrapTextAtRightBorder {
			get { return wrapTextAtRightBorder; }
			set { wrapTextAtRightBorder = value; }
		}

        private CellHorizontalAlignment horizontalAlignment = CellHorizontalAlignment.General;

		public CellHorizontalAlignment HorizontalAlignment {
			get { return horizontalAlignment; }
			set { horizontalAlignment = value; }
		}

        private CellVerticalAlignment verticalAlignment = CellVerticalAlignment.Top;

		public CellVerticalAlignment VerticalAlignment {
			get { return verticalAlignment; }
			set { verticalAlignment = value; }
		}

        private bool justifyLastLine = false;

		public bool JustifyTextAtLastLine {
			get { return justifyLastLine; }
			set { justifyLastLine = value; }
		}

        private byte rotation = 0;

        /// <summary>
        /// 0 Not rotated
		/// 1-90 1 to 90 degrees counterclockwise
		/// 91-180 1 to 90 degrees clockwise
		/// 255 Letters are stacked top-to-bottom, but not rotated
        /// </summary>
		public byte Rotation {
			get { return rotation; }
			set { rotation = value; }
		}
        #endregion

        #region Borders
        private BorderLineStyle topBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle TopBorderLineStyle {
			get { return topBorderLineStyle; }
			set { topBorderLineStyle = value; }
		}

        private Color topBorderLineColor = Color.Black;

		public Color TopBorderLineColor {
			get { return topBorderLineColor; }
			set { topBorderLineColor = value; }
		}

        private BorderLineStyle leftBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle LeftBorderLineStyle {
			get { return leftBorderLineStyle; }
			set { leftBorderLineStyle = value; }
		}

        private Color leftBorderLineColor = Color.Black;

		public Color LeftBorderLineColor {
			get { return leftBorderLineColor; }
			set { leftBorderLineColor = value; }
		}

        private BorderLineStyle rightBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle RightBorderLineStyle {
			get { return rightBorderLineStyle; }
			set { rightBorderLineStyle = value; }
		}

        private Color rightBorderLineColor = Color.Black;

		public Color RightBorderLineColor {
			get { return rightBorderLineColor; }
			set { rightBorderLineColor = value; }
		}

        private BorderLineStyle bottomBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle BottomBorderLineStyle {
			get { return bottomBorderLineStyle; }
			set { bottomBorderLineStyle = value; }
		}

        private Color bottomBorderLineColor = Color.Black;

		public Color BottomBorderLineColor {
			get { return bottomBorderLineColor; }
			set { bottomBorderLineColor = value; }
		}
        #endregion

        #region Background
        private Color backgroundColor = Color.White;

		public Color BackgroundColor {
			get { return backgroundColor; }
			set { backgroundColor = value; }
		}

        private Color backgroundPatternColor = Color.White;

		public Color BackgroundPatternColor {
			get { return backgroundPatternColor; }
			set { backgroundPatternColor = value; }
		}

        private CellBackgroundPattern backgroundPattern = CellBackgroundPattern.None;

		public CellBackgroundPattern BackgroundPattern {
			get { return backgroundPattern; }
			set { backgroundPattern = value; }
		}
        #endregion
    }
}
