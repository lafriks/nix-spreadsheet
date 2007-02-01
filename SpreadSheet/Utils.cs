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
