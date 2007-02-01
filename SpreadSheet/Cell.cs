using System;

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Summary description for Cell.
	/// </summary>
    public class Cell : AbstractCell
    {
        // Row of the value
        private int row;
        // Column of the value
        private int column;

        public Cell (int row, int col)
        {
            this.column = col;
            this.row = row;
        }

        public Cell (string name)
        {
            Utils.ParseCellName(name, ref this.row, ref this.column);
        }

        internal int RowIndex
        {
            get
            {
                return this.row;
            }
        }

        internal int ColumnIndex
        {
            get
            {
                return this.column;
            }
        }

        internal string Name
        {
            get
            {
                return Utils.CellName(this.column, this.row);
            }
        }
    }
}
