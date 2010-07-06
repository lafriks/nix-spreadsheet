using System;
using System.Collections.Generic;
using System.Text;

namespace Nix.SpreadSheet
{
    public class ColumnRange
    {
        private int startColumn;

        public int StartColumn
        {
            get { return startColumn; }
        }

        private int endColumn;

        public int EndColumn
        {
            get { return endColumn; }
        }

        private Sheet sheet;

        public Sheet Sheet
        {
            get { return sheet; }
        }

        public ColumnRange(Sheet sheet, int startColumn, int endColumn)
        {
            this.sheet = sheet;
            this.startColumn = startColumn;
            this.endColumn = endColumn;
        }

        public uint Width
        {
            set
            {
                for (int i = this.StartColumn; i <= this.EndColumn; i++)
                {
                    this.sheet.Columns[i].Width = value;
                }
            }
            get
            {
                uint width = this.Sheet.Columns[this.StartColumn].Width;
                for (int i = this.StartColumn + 1; i <= this.EndColumn; i++)
                {
                    if (this.sheet.Columns[i].Width != width)
                        throw new Exception("ColumnRange columns does not have same width");
                }
                return width;
            }
        }
    }
}
