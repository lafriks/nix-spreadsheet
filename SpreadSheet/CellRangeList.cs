using System;
using System.Collections.Generic;
using System.Text;

namespace Nix.SpreadSheet
{
    public class CellRangeList : List<CellRange>
    {
        public bool IntersectsWith(CellRange range)
        {
            foreach (CellRange cr in this)
            {
                if (Math.Max(cr.FirstColumn, range.FirstColumn) <= Math.Min(cr.LastColumn, range.LastColumn) && Math.Max(cr.FirstRow, range.FirstRow) <= Math.Min(cr.LastRow, range.LastRow))
                    return true;
            }
            return false;
        }

        public CellRange GetAtPosition(int row, int column)
        {
            foreach (CellRange cr in this)
            {
                if (cr.FirstColumn <= column && cr.FirstRow <= row && cr.LastColumn >= column && cr.LastRow >= row)
                    return cr;
            }
            return null;
        }
    }
}
