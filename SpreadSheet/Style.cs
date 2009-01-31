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
    public class Style : IEquatable<Style>
    {
        private static Style def = new Style();

        public static Style Default
        {
            get
            {
                return def;
            }
            set
            {
                def = value;
            }
        }

        #region Cell protection
        private bool cellLocked = false;

        public bool CellLocked {
            get
            {
                return this.cellLocked;
            }
            set
            {
                this.cellLocked = value;
            }
        }

        private bool hiddenFormula = false;

        public bool HiddenFormula {
            get
            {
                return this.hiddenFormula;
            }
            set
            {
                this.hiddenFormula = value;
            }
        }
        #endregion

        #region Text properties
        private bool wrapText = false;

        /// <summary>
        /// Gets or sets if cell value is wrapped.
        /// </summary>
        public bool WrapText {
            get
            {
                return this.wrapText;
            }
            set
            {
                this.wrapText = value;
            }
        }

        private bool shrinkToFit = false;

        /// <summary>
        /// Gets or sets if cell value should shrunk to fit the cell.
        /// </summary>
        public bool ShrinkToFit {
            get
            {
                return this.shrinkToFit;
            }
            set
            {
                this.shrinkToFit = value;
            }
        }
        #endregion

        #region Font
        private Font font = new Font();

        public Font Font {
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
            get
            {
                return this.format;
            }
            set
            {
                this.format = value;
            }
        }
        #endregion

        #region Indent level, text direction
        private byte indentLevel = 0;

        public byte IndentLevel {
            get
            {
                return this.indentLevel;
            }
            set
            {
                this.indentLevel = value;
            }
        }

        private TextDirection textDirection = TextDirection.Automatic;

        public TextDirection TextDirection {
            get
            {
                return this.textDirection;
            }
            set
            {
                this.textDirection = value;
            }
        }
        #endregion

        #region Alignment
        private bool wrapTextAtRightBorder = true;

        public bool WrapTextAtRightBorder {
            get
            {
                return this.wrapTextAtRightBorder;
            }
            set
            {
                this.wrapTextAtRightBorder = value;
            }
        }

        private CellHorizontalAlignment horizontalAlignment = CellHorizontalAlignment.General;

        public CellHorizontalAlignment HorizontalAlignment {
            get
            {
                return this.horizontalAlignment;
            }
            set
            {
                this.horizontalAlignment = value;
            }
        }

        private CellVerticalAlignment verticalAlignment = CellVerticalAlignment.Top;

        public CellVerticalAlignment VerticalAlignment {
            get
            {
                return this.verticalAlignment;
            }
            set
            {
                this.verticalAlignment = value;
            }
        }

        private bool justifyTextAtLastLine = false;

        public bool JustifyTextAtLastLine {
            get
            {
                return this.justifyTextAtLastLine;
            }
            set
            {
                this.justifyTextAtLastLine = value;
            }
        }

        private byte rotation = 0;

        /// <summary>
        /// 0 Not rotated;
        /// 1-90 1 to 90 degrees counterclockwise;
        /// 91-180 1 to 90;
        /// degrees clockwise;
        /// 255 Letters are stacked top-to-bottom, but not rotated.
        /// </summary>
        public byte Rotation {
            get
            {
                return this.rotation;
            }
            set
            {
                this.rotation = value;
            }
        }
        #endregion

        #region Borders
        private BorderLineStyle topBorderLineStyle = BorderLineStyle.None;

        public BorderLineStyle TopBorderLineStyle {
            get
            {
                return this.topBorderLineStyle;
            }
            set
            {
                this.topBorderLineStyle = value;
            }
        }

        private Color topBorderLineColor = Color.Black;

        public Color TopBorderLineColor {
            get
            {
                return this.topBorderLineColor;
            }
            set
            {
                this.topBorderLineColor = value;
            }
        }

        private BorderLineStyle leftBorderLineStyle = BorderLineStyle.None;

        public BorderLineStyle LeftBorderLineStyle {
            get
            {
                return this.leftBorderLineStyle;
            }
            set
            {
                this.leftBorderLineStyle = value;
            }
        }

        private Color leftBorderLineColor = Color.Black;

        public Color LeftBorderLineColor {
            get
            {
                return this.leftBorderLineColor;
            }
            set
            {
                this.leftBorderLineColor = value;
            }
        }

        private BorderLineStyle rightBorderLineStyle = BorderLineStyle.None;

        public BorderLineStyle RightBorderLineStyle {
            get
            {
                return this.rightBorderLineStyle;
            }
            set
            {
                this.rightBorderLineStyle = value;
            }
        }

        private Color rightBorderLineColor = Color.Black;

        public Color RightBorderLineColor {
            get
            {
                return this.rightBorderLineColor;
            }
            set
            {
                this.rightBorderLineColor = value;
            }
        }

        private BorderLineStyle bottomBorderLineStyle = BorderLineStyle.None;

        public BorderLineStyle BottomBorderLineStyle {
            get
            {
                return this.bottomBorderLineStyle;
            }
            set
            {
                this.bottomBorderLineStyle = value;
            }
        }

        private Color bottomBorderLineColor = Color.Black;

        public Color BottomBorderLineColor {
            get
            {
                return this.bottomBorderLineColor;
            }
            set
            {
                this.bottomBorderLineColor = value;
            }
        }
        #endregion

        #region Background
        private Color backgroundColor = Color.White;

        public Color BackgroundColor {
            get
            {
                return this.backgroundColor;
            }
            set
            {
                this.backgroundColor = value;
            }
        }

        private Color backgroundPatternColor = Color.White;

        public Color BackgroundPatternColor {
            get
            {
                return this.backgroundPatternColor;
            }
            set
            {
                this.backgroundPatternColor = value;
            }
        }

        private CellBackgroundPattern backgroundPattern = CellBackgroundPattern.None;

        public CellBackgroundPattern BackgroundPattern {
            get
            {
                return this.backgroundPattern;
            }
            set
            {
                this.backgroundPattern = value;
            }
        }
        #endregion

        public bool Equals(Style other)
        {
            return ! ((this.CellLocked != other.CellLocked)
                     || (this.HiddenFormula != other.HiddenFormula)
                     || (this.WrapText != other.WrapText)
                     || (this.ShrinkToFit != other.ShrinkToFit)
                     || (this.Font.Equals(other.Font) == false)
                     || (this.Format != other.Format)
                     || (this.IndentLevel != other.IndentLevel)
                     || (this.TextDirection != other.TextDirection)
                     || (this.WrapTextAtRightBorder != other.WrapTextAtRightBorder)
                     || (this.HorizontalAlignment != other.HorizontalAlignment)
                     || (this.VerticalAlignment != other.VerticalAlignment)
                     || (this.JustifyTextAtLastLine != other.JustifyTextAtLastLine)
                     || (this.Rotation != other.Rotation)
                     || (this.TopBorderLineStyle != other.TopBorderLineStyle)
                     || (this.TopBorderLineColor != other.TopBorderLineColor)
                     || (this.LeftBorderLineStyle != other.LeftBorderLineStyle)
                     || (this.LeftBorderLineColor != other.LeftBorderLineColor)
                     || (this.RightBorderLineStyle != other.RightBorderLineStyle)
                     || (this.RightBorderLineColor != other.RightBorderLineColor)
                     || (this.BottomBorderLineStyle != other.BottomBorderLineStyle)
                     || (this.BottomBorderLineColor != other.BottomBorderLineColor)
                     || (this.BackgroundColor != other.BackgroundColor)
                     || (this.BackgroundPatternColor != other.BackgroundPatternColor)
                     || (this.BackgroundPattern != other.BackgroundPattern));
        }
    }
}
