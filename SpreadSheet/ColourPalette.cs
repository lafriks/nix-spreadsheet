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
    /// Summary description for Cell.
    /// </summary>
    public class ColourPalette
    {
        public ColourPalette()
        {
        }

        public ushort GetColourIndex(Color col)
        {
            ushort result = 0;
            bool colorFound = true;

            switch (col.ToKnownColor())
	        {
                case KnownColor.Black:
                    result = 0;
                    break;
                case KnownColor.White:
                    result = 1;
                    break;
                case KnownColor.Red:
                    result = 2;
                    break;
                case KnownColor.Green:
                    result = 3;
                    break;
                case KnownColor.Blue:
                    result = 4;
                    break;
                case KnownColor.Yellow:
                    result = 5;
                    break;
                case KnownColor.Magenta:
                    result = 6;
                    break;
                case KnownColor.Cyan:
                    result = 7;
                    break;
		        default:
                    colorFound = false;
                    break;
	        }

            return result;
        }
    }
}
