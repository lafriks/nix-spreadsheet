/*
 * Library for creating Microsoft Excel files. * Copyright (C) 2007, Lauris Bukšis-Haberkorns <lauris@nix.lv> *
 * This library is free software; you can redistribute it and/or * modify it under the terms of the GNU Lesser General Public * License as published by the Free Software Foundation; either * version 2.1 of the License, or (at your option) any later version. *
 * This library is distributed in the hope that it will be useful, * but WITHOUT ANY WARRANTY; without even the implied warranty of * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU * Lesser General Public License for more details. *
 * You should have received a copy of the GNU Lesser General Public * License along with this library; if not, write to the Free Software * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA */

using System;
using System.Drawing;

namespace Nix.SpreadSheet
{
	public class Font
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
        
        internal bool Equals(Font obj)
        {
            return ! (this.color != obj.color || this.italic != obj.italic
                      || this.weight != obj.weight || this.strikeout != obj.strikeout
                      || this.name != obj.name || this.size != obj.size
                      || this.underline != obj.underline || this.scriptpos != obj.scriptpos );
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
        private int weight = NormalWeight;
        public int Weight
        {
            get
            {
                return this.weight;
            }
            set
            {
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
        private int size = 200;
        public int Size
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
}
