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
    /// Cell.
    /// </summary>
    public class Cell : AbstractCell
    {
        // Row of the value
        private int row;
        // Column of the value
        private int column;

        private Sheet sheet;

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// </summary>
        /// <param name="row">Row position</param>
        /// <param name="col">Column position</param>
        internal Cell (Sheet sheet, int row, int col)
        {
            this.sheet = sheet;
            this.column = col;
            this.row = row;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// </summary>
        /// <param name="name">Cell name.</param>
        public Cell (string name)
        {
            Utils.ParseCellName(name, out this.row, out this.column);
        }

        /// <summary>
        /// Cell row index.
        /// </summary>
        public int RowIndex
        {
            get
            {
                return this.row;
            }
        }

        /// <summary>
        /// Cell column index.
        /// </summary>
        public int ColumnIndex
        {
            get
            {
                return this.column;
            }
        }

        /// <summary>
        /// Cell name.
        /// </summary>
        public string Name
        {
            get
            {
                return Utils.CellName(this.row, this.column);
            }
        }


        #region Value
        internal object InternalValue
        {
            get
            {
                return base.Value;
            }
        }

        internal string InternalDisplayValue
        {
            get
            {
                return base.DisplayValue;
            }
        }

        /// <summary>
        /// Gets or sets value. When setting value formula is removed.
        /// </summary>
        public override object Value
        {
            get
            {
                CellRange cr = this.sheet.mergedCells.GetAtPosition(row, column);
                if (cr != null && (row != cr.FirstRow || column != cr.FirstColumn))
                {
                    if (this.sheet.Document.MergedCellsBehaviour == MergedCellsBehaviour.AccessFirstCell)
                    {
                        return this.sheet[cr.FirstRow, cr.FirstColumn].Value;
                    }
                    else if (this.sheet.Document.MergedCellsBehaviour == MergedCellsBehaviour.ThrowExceptionOnAccess
                                && (row != cr.FirstRow || column != cr.FirstColumn))
                    {
                        throw new Exception("Can not access merged cell value at this position");
                    }
                }

                return base.Value;
            }
            set
            {
                CellRange cr = this.sheet.mergedCells.GetAtPosition(row, column);
                if (cr != null && (row != cr.FirstRow || column != cr.FirstColumn))
                {
                    if (this.sheet.Document.MergedCellsBehaviour == MergedCellsBehaviour.AccessFirstCell)
                    {
                        this.sheet[cr.FirstRow, cr.FirstColumn].Value = value;
                    }
                    else if (this.sheet.Document.MergedCellsBehaviour == MergedCellsBehaviour.ThrowExceptionOnAccess
                                && (row != cr.FirstRow || column != cr.FirstColumn))
                    {
                        throw new Exception("Can not access merged cell value at this position");
                    }
                }
                else
                {
                    base.Value = value;
                }
            }
        }

        /// <summary>
        /// Gets display value
        /// </summary>
        public override string DisplayValue
        {
            get
            {
                CellRange cr = this.sheet.mergedCells.GetAtPosition(row, column);
                if (cr != null && (row != cr.FirstRow || column != cr.FirstColumn))
                {
                    if (this.sheet.Document.MergedCellsBehaviour == MergedCellsBehaviour.AccessFirstCell)
                    {
                        return this.sheet[cr.FirstRow, cr.FirstColumn].DisplayValue;
                    }
                    else if (this.sheet.Document.MergedCellsBehaviour == MergedCellsBehaviour.ThrowExceptionOnAccess
                                && (row != cr.FirstRow || column != cr.FirstColumn))
                    {
                        throw new Exception("Can not access merged cell value at this position");
                    }
                }

                return base.DisplayValue;
            }
        }
        #endregion

        #region Cell height calculation
        public ushort CalculateCellHeight()
        {
            CellRange cr = this.sheet.mergedCells.GetAtPosition(this.RowIndex, this.ColumnIndex);
            int width;
            int height;
            string val;
            bool wrap;
            if (cr != null)
            {
                // Merged cell
                height = 0;
                for (int r = cr.FirstRow; r <= cr.LastRow; ++r)
                    height += (this.sheet[this.RowIndex].Height ?? 17);
                width = 0;
                for (int c = cr.FirstColumn; c <= cr.LastColumn; ++c)
                    width += (int)this.sheet.Columns[c].Width;

                Cell cell = this.sheet[cr.FirstRow, cr.FirstColumn];
                val = cell.DisplayValue;
                wrap = (cell.formatting != null ? cell.formatting.WrapTextAtRightBorder : false);
            }
            else
            {
                width = (int)this.sheet.Columns[this.ColumnIndex].Width;
                height = this.sheet[this.RowIndex].Height ?? 17;
                val = this.DisplayValue;
                wrap = (this.formatting != null ? this.formatting.WrapTextAtRightBorder : false);
            }

            // TODO: Bring back word wrap measurment
            //TextFormatFlags tff = TextFormatFlags.Default
            //                      | TextFormatFlags.ExpandTabs
            //                      | TextFormatFlags.NoPrefix
            //                      | TextFormatFlags.TextBoxControl
            //                      | TextFormatFlags.NoPadding
            //                      | TextFormatFlags.NoFullWidthCharacterBreak;
            //if (wrap)
            //    tff |= TextFormatFlags.WordBreak;

            System.Drawing.Font font = (this.HasFormatting ? this.Formatting.Font : Style.Default.Font).ToNativeFont();

            var s = Graphics.FromImage(new Bitmap(1, 1)).MeasureString(val, font, width + 6);
            return (ushort)Math.Round(s.Height);
        }
        #endregion
    }
}
