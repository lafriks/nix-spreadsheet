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
	public class Font : IEquatable<Font>, ICloneable
	{
        public const int NormalWeight = 400;

        public const int BoldWeight = 700;

        public const int MaxWeight = 1000;

        public const int MinWeight = 100;

        #region Default and equals method
        private static Font def = null;
        public static Font Default
        {
            get
            {
                if (def == null)
                    def = new Font();
                return def;
            }
        }

        public bool Equals(Font other)
        {
            return ! (this.color != other.color || this.italic != other.italic
                      || this.weight != other.weight || this.strikeout != other.strikeout
                      || this.name != other.name || this.size != other.size
                      || this.underline != other.underline || this.scriptpos != other.scriptpos
                      || this.fontFace != other.fontFace || this.charSet != other.charSet);
        }

		public object Clone()
		{
			Font font = new Font();
			font.charSet = this.charSet;
			font.color = this.color;
			font.fontFace = this.fontFace;
			font.italic = this.italic;
			font.name = this.name;
			font.scriptpos = this.scriptpos;
			font.size = this.size;
			font.strikeout = this.strikeout;
			font.underline = this.underline;
			font.weight = this.weight;
			return font;
		}
		#endregion

        #region Color
        private Color color = Color.Black;

        public Color Color
        {
            get
            {
                return this.color;
            }
            set
            {
                this.color = value;
            }
        }
        #endregion
        
        #region Italic, strikeout etc
        private bool italic = false;
        public bool Italic
        {
            get
            {
                return this.italic;
            }
            set
            {
                this.italic = value;
            }
        }

        private bool strikeout = false;
        public bool Strikeout
        {
            get
            {
                return this.strikeout;
            }
            set
            {
                this.strikeout = value;
            }
        }
        #endregion

        #region Weight
        private ushort weight = NormalWeight;
        public ushort Weight
        {
            get
            {
                return this.weight;
            }
            set
            {
            	if ( value < MinWeight || value > MaxWeight )
            		throw new ArgumentOutOfRangeException("Weight");
                this.weight = value;
            }
        }
        #endregion

        #region Name
        private string name = "Arial";
        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }
        #endregion

        #region Size
        private ushort size = 200;
        public ushort Size
        {
            get
            {
                return this.size;
            }
            set
            {
                this.size = value;
            }
        }
        #endregion

        #region Underline style
        private UnderlineStyle underline = UnderlineStyle.None;
        public UnderlineStyle UnderlineStyle
        {
            get
            {
                return this.underline;
            }
            set
            {
                this.underline = value;
            }
        }
        #endregion

        #region Script position
        private ScriptPosition scriptpos = ScriptPosition.Normal;
        public ScriptPosition ScriptPosition
        {
            get
            {
                return this.scriptpos;
            }
            set
            {
                this.scriptpos = value;
            }
        }
        #endregion

        #region Font face
        private FontFace fontFace = FontFace.None;
        
		public FontFace FontFace {
			get { return fontFace; }
			set { fontFace = value; }
		}
        #endregion

		#region Char set        
        private CharSet charSet = CharSet.Default;
        
		public CharSet CharSet {
			get { return charSet; }
			set { charSet = value; }
		}
		#endregion
    }

    public enum UnderlineStyle
    {
          None,
          Single,
          Double,
          SingleAccounting,
          DoubleAccounting
    }

    public enum ScriptPosition
    {
          Normal,
          Superscript,
          Subscript
    }

    public enum FontFace
    {
    	None,
    	Modern,
    	Roman,
    	Swiss,
    	Script,
    	Decorative
    }

    public enum CharSet
    {
		Ansi,
		Arabic,
		Baltic,
		ChineseBig5,
		Default,
		EastEurope,
		GB2312,
		Greek,
		Hangeul,
		Hebrew,
		Johab,
		MAC,
		OEM,
		Russian,
		Shiftjis,
		Symbol,
		Thai,
		Turkish
    }
}
