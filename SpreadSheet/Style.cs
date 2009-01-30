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
	public class Style
	{
		private uint modified = 0;

		#region Cell protection
		private bool cellLocked = false;

		public bool CellLocked
		{
			get
			{
				return this.cellLocked;
			}
			set
			{
				if ( this.cellLocked != value )
				{
					this.modified |= 1;
					this.cellLocked = value;
				}
			}
		}

		public bool IsModifiedCellLocked ()
		{
			return ((this.modified & 1) == 1);
		}

		private bool hiddenFormula = false;

		public bool HiddenFormula
		{
			get
			{
				return this.hiddenFormula;
			}
			set
			{
				if ( this.hiddenFormula != value )
				{
					this.modified |= 2;
					this.hiddenFormula = value;
				}
			}
		}

		public bool IsModifiedHiddenFormula ()
		{
			return ((this.modified & 2) == 2);
		}
		#endregion

		#region Text properties
		private bool wrapText = false;

		/// <summary>
		/// Gets or sets if cell value is wrapped.
		/// </summary>
		public bool WrapText
		{
			get
			{
				return this.wrapText;
			}
			set
			{
				if ( this.wrapText != value )
				{
					this.modified |= 4;
					this.wrapText = value;
				}
			}
		}

		public bool IsModifiedWrapText ()
		{
			return ((this.modified & 4) == 4);
		}

		private bool shrinkToFit = false;

		/// <summary>
		/// Gets or sets if cell value should shrunk to fit the cell.
		/// </summary>
		public bool ShrinkToFit
		{
			get
			{
				return this.shrinkToFit;
			}
			set
			{
				if ( this.shrinkToFit != value )
				{
					this.modified |= 8;
					this.shrinkToFit = value;
				}
			}
		}

		public bool IsModifiedShrinkToFit ()
		{
			return ((this.modified & 8) == 8);
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
				if ( this.font != value )
				{
					this.modified |= 16;
					this.font = value;
				}
			}
		}

		public bool IsModifiedFont ()
		{
			return ((this.modified & 16) == 16);
		}
		#endregion

		#region Format
		private string format = string.Empty;

		public string Format
		{
			get
			{
				return this.format;
			}
			set
			{
				if ( this.format != value )
				{
					this.modified |= 32;
					this.format = value;
				}
			}
		}

		public bool IsModifiedFormat ()
		{
			return ((this.modified & 32) == 32);
		}
		#endregion

		#region Alignment
		private bool wrapTextAtRightBorder = true;

		public bool WrapTextAtRightBorder
		{
			get
			{
				return this.wrapTextAtRightBorder;
			}
			set
			{
				if ( this.wrapTextAtRightBorder != value )
				{
					this.modified |= 64;
					this.wrapTextAtRightBorder = value;
				}
			}
		}

		public bool IsModifiedWrapTextAtRightBorder ()
		{
			return ((this.modified & 64) == 64);
		}

		private CellHorizontalAlignment horizontalAlignment = CellHorizontalAlignment.General;

		public CellHorizontalAlignment HorizontalAlignment
		{
			get
			{
				return this.horizontalAlignment;
			}
			set
			{
				if ( this.horizontalAlignment != value )
				{
					this.modified |= 128;
					this.horizontalAlignment = value;
				}
			}
		}

		public bool IsModifiedHorizontalAlignment ()
		{
			return ((this.modified & 128) == 128);
		}

		private CellVerticalAlignment verticalAlignment = CellVerticalAlignment.Top;

		public CellVerticalAlignment VerticalAlignment
		{
			get
			{
				return this.verticalAlignment;
			}
			set
			{
				if ( this.verticalAlignment != value )
				{
					this.modified |= 256;
					this.verticalAlignment = value;
				}
			}
		}

		public bool IsModifiedVerticalAlignment ()
		{
			return ((this.modified & 256) == 256);
		}

		private bool justifyTextAtLastLine = false;

		public bool JustifyTextAtLastLine
		{
			get
			{
				return this.justifyTextAtLastLine;
			}
			set
			{
				if ( this.justifyTextAtLastLine != value )
				{
					this.modified |= 512;
					this.justifyTextAtLastLine = value;
				}
			}
		}

		public bool IsModifiedJustifyTextAtLastLine ()
		{
			return ((this.modified & 512) == 512);
		}

		private byte rotation = 0;

		/// <summary>
		/// 0 Not rotated;
		/// 1-90 1 to 90 degrees counterclockwise;
		/// 91-180 1 to 90;
		/// degrees clockwise;
		/// 255 Letters are stacked top-to-bottom, but not rotated.
		/// </summary>
		public byte Rotation
		{
			get
			{
				return this.rotation;
			}
			set
			{
				if ( this.rotation != value )
				{
					this.modified |= 1024;
					this.rotation = value;
				}
			}
		}

		public bool IsModifiedRotation ()
		{
			return ((this.modified & 1024) == 1024);
		}
		#endregion

		#region Borders
		private BorderLineStyle topBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle TopBorderLineStyle
		{
			get
			{
				return this.topBorderLineStyle;
			}
			set
			{
				if ( this.topBorderLineStyle != value )
				{
					this.modified |= 2048;
					this.topBorderLineStyle = value;
				}
			}
		}

		public bool IsModifiedTopBorderLineStyle ()
		{
			return ((this.modified & 2048) == 2048);
		}

		private Color topBorderLineColor = Color.Black;

		public Color TopBorderLineColor
		{
			get
			{
				return this.topBorderLineColor;
			}
			set
			{
				if ( this.topBorderLineColor != value )
				{
					this.modified |= 4096;
					this.topBorderLineColor = value;
				}
			}
		}

		public bool IsModifiedTopBorderLineColor ()
		{
			return ((this.modified & 4096) == 4096);
		}

		private BorderLineStyle leftBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle LeftBorderLineStyle
		{
			get
			{
				return this.leftBorderLineStyle;
			}
			set
			{
				if ( this.leftBorderLineStyle != value )
				{
					this.modified |= 8192;
					this.leftBorderLineStyle = value;
				}
			}
		}

		public bool IsModifiedLeftBorderLineStyle ()
		{
			return ((this.modified & 8192) == 8192);
		}

		private Color leftBorderLineColor = Color.Black;

		public Color LeftBorderLineColor
		{
			get
			{
				return this.leftBorderLineColor;
			}
			set
			{
				if ( this.leftBorderLineColor != value )
				{
					this.modified |= 16384;
					this.leftBorderLineColor = value;
				}
			}
		}

		public bool IsModifiedLeftBorderLineColor ()
		{
			return ((this.modified & 16384) == 16384);
		}

		private BorderLineStyle rightBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle RightBorderLineStyle
		{
			get
			{
				return this.rightBorderLineStyle;
			}
			set
			{
				if ( this.rightBorderLineStyle != value )
				{
					this.modified |= 32768;
					this.rightBorderLineStyle = value;
				}
			}
		}

		public bool IsModifiedRightBorderLineStyle ()
		{
			return ((this.modified & 32768) == 32768);
		}

		private Color rightBorderLineColor = Color.Black;

		public Color RightBorderLineColor
		{
			get
			{
				return this.rightBorderLineColor;
			}
			set
			{
				if ( this.rightBorderLineColor != value )
				{
					this.modified |= 65536;
					this.rightBorderLineColor = value;
				}
			}
		}

		public bool IsModifiedRightBorderLineColor ()
		{
			return ((this.modified & 65536) == 65536);
		}

		private BorderLineStyle bottomBorderLineStyle = BorderLineStyle.None;

		public BorderLineStyle BottomBorderLineStyle
		{
			get
			{
				return this.bottomBorderLineStyle;
			}
			set
			{
				if ( this.bottomBorderLineStyle != value )
				{
					this.modified |= 131072;
					this.bottomBorderLineStyle = value;
				}
			}
		}

		public bool IsModifiedBottomBorderLineStyle ()
		{
			return ((this.modified & 131072) == 131072);
		}

		private Color bottomBorderLineColor = Color.Black;

		public Color BottomBorderLineColor
		{
			get
			{
				return this.bottomBorderLineColor;
			}
			set
			{
				if ( this.bottomBorderLineColor != value )
				{
					this.modified |= 262144;
					this.bottomBorderLineColor = value;
				}
			}
		}

		public bool IsModifiedBottomBorderLineColor ()
		{
			return ((this.modified & 262144) == 262144);
		}
		#endregion

		#region Background
		private Color backgroundColor = Color.White;

		public Color BackgroundColor
		{
			get
			{
				return this.backgroundColor;
			}
			set
			{
				if ( this.backgroundColor != value )
				{
					this.modified |= 524288;
					this.backgroundColor = value;
				}
			}
		}

		public bool IsModifiedBackgroundColor ()
		{
			return ((this.modified & 524288) == 524288);
		}

		private Color backgroundPatternColor = Color.White;

		public Color BackgroundPatternColor
		{
			get
			{
				return this.backgroundPatternColor;
			}
			set
			{
				if ( this.backgroundPatternColor != value )
				{
					this.modified |= 1048576;
					this.backgroundPatternColor = value;
				}
			}
		}

		public bool IsModifiedBackgroundPatternColor ()
		{
			return ((this.modified & 1048576) == 1048576);
		}

		private CellBackgroundPattern backgroundPattern = CellBackgroundPattern.None;

		public CellBackgroundPattern BackgroundPattern
		{
			get
			{
				return this.backgroundPattern;
			}
			set
			{
				if ( this.backgroundPattern != value )
				{
					this.modified |= 2097152;
					this.backgroundPattern = value;
				}
			}
		}

		public bool IsModifiedBackgroundPattern ()
		{
			return ((this.modified & 2097152) == 2097152);
		}
		#endregion
	}
}
