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
using System.Collections;

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Summary description for Utils.
    /// </summary>
    internal class Utils
    {
        #region Cell name utils
        private static string colChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static string CellName(int row, int col)
        {
            string name = string.Empty;

            while (col >= 26)
            {
                name = colChars[col % 26] + name;
                col = (col / 26) - 1;
            }
            name = colChars[col] + name;

            return name + row.ToString();
        }

        public static void ParseCellName(string name, ref int row, ref int col)
        {
            col = 0;
            row = 0;
            int pos = -1;
            int pow = 0;

            foreach (char a in name)
            {
                if (colChars.IndexOf(a) > -1)
                {
                    if (row > 0)
                        throw new ArgumentException("Invalid cell name");
                    pos++;
                }
                else
                {
                    while (pos >= 0)
                    {
                        col += (colChars.IndexOf(name[pos--]) + 1) * (int)Math.Pow(26, pow++);
                    }
                    row = row * 10 + (a - 48);
                }
            }
            if (row == 0 || col == 0)
                throw new ArgumentException("Invalid cell name");
            col --;
            row --;
        }
        #endregion
    }
}
