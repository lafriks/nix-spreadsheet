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
using System.Collections.Generic;

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Summary description for Cell.
    /// </summary>
    public class ColorPalette : IEnumerable<uint>
    {

        private SortedDictionary<ushort, uint> m_colors = new SortedDictionary<ushort, uint>();
        private SortedDictionary<uint, ushort> m_indexes = new SortedDictionary<uint, ushort>();
        
        private static uint maxIdx = 63;

        public ushort GetColorIndex(Color col)
        {
            ushort result = 0;
            bool colorFound = false;

            /*switch (col.ToKnownColor())
	        {
                case KnownColor.Black:
                    result = 0x08;
                    break;
                case KnownColor.White:
                    result = 0x09;
                    break;
                case KnownColor.Red:
                    result = 0x0A;
                    break;
                case KnownColor.Green:
                    result = 0x11;
                    break;
                case KnownColor.Blue:
                    result = 0x0C;
                    break;
                case KnownColor.Yellow:
                    result = 0x0D;
                    break;
                case KnownColor.Fuchsia:
                case KnownColor.Magenta:
                    result = 0x0E;
                    break;
                case KnownColor.Aqua:
                case KnownColor.Cyan:
                    result = 0x0F;
                    break;
                case KnownColor.Brown:
                    result = 0x10;
                    break;
                case KnownColor.Gray:
                    result = 0x17;
                    break;
                case KnownColor.Lime:
                    result = 0x0B;
                    break;
                case KnownColor.Orange:
                    result = 0x35;
                    break;
                case KnownColor.Purple:
                    result = 0x14;
                    break;
                case KnownColor.Silver:
                    result = 0x16;
                    break;
		        default:
                    colorFound = false;
                    break;
	        }*/
            
            if (!colorFound)
            {
            	uint color = (((uint)col.B)<<16)|(((uint)col.G)<<8)|((uint)col.R);
            	if ( this.m_indexes.ContainsKey(color) )
            	{
                	result = this.m_indexes[color];
                	colorFound =true;
            	}
            	else
            	{
            		result = (ushort)(m_indexes.Count + 8);
            		if (result > maxIdx)
            		{
						throw new IndexOutOfRangeException("Colors' count exceeded");
            		}
            		else
            		{
            			m_indexes.Add(color, result);
            			m_colors.Add(result, color);
            		}
            	}
            }

            return result;
        }
        
        public uint this[ushort idx]
		{
			get { return m_colors[idx]; }
		}
        
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.m_colors.Values.GetEnumerator();
		}
		
		IEnumerator<uint> IEnumerable<uint>.GetEnumerator()
		{
			return this.m_colors.Values.GetEnumerator();
		}
		
		public ushort Count
		{
			get { return (ushort)m_colors.Count;}
		}
    }
}
