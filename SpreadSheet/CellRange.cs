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
	/// Description of CellRange.
	/// </summary>
	public class CellRange
	{
		private Sheet sheet;
		/// <summary>
		/// Sheet.
		/// </summary>
		public Sheet Sheet
		{
			get { return sheet; }
		}

		private int firstRow;
		/// <summary>
		/// First row in range.
		/// </summary>
		public int FirstRow {
			get { return firstRow; }
		}

		private int firstColumn;
		/// <summary>
		/// First columnt in range.
		/// </summary>
		public int FirstColumn {
			get { return firstColumn; }
		}

		private int lastRow;
		/// <summary>
		/// Last row in range.
		/// </summary>
		public int LastRow {
			get { return lastRow; }
		}

		private int lastColumn;
		/// <summary>
		/// Last column in range.
		/// </summary>
		public int LastColumn {
			get { return lastColumn; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="CellRange"/> class.
		/// </summary>
		/// <param name="sheet">Sheet.</param>
		/// <param name="firstRow">First row.</param>
		/// <param name="firstColumn">First column.</param>
		/// <param name="lastRow">Last row.</param>
		/// <param name="lastColumn">Last column.</param>
		public CellRange(Sheet sheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
		{
			this.sheet = sheet;
			this.firstRow = firstRow;
			this.firstColumn = firstColumn;
			this.lastRow = lastRow;
			this.lastColumn = lastColumn;
		}

		/// <summary>
		/// Draw border around range of cells.
		/// </summary>
		/// <param name="borderColor">Border color.</param>
		/// <param name="borderLineStyle">Border line style.</param>
		/// <returns>Current range instance.</returns>
		public CellRange DrawBorder(Color borderColor, BorderLineStyle borderLineStyle)
		{
			for (int i = FirstRow; i <= LastRow; i++)
			{
				sheet[i, FirstColumn].Formatting.LeftBorderLineColor = borderColor;
				sheet[i, FirstColumn].Formatting.LeftBorderLineStyle = borderLineStyle;
				sheet[i, LastColumn].Formatting.RightBorderLineColor = borderColor;
				sheet[i, LastColumn].Formatting.RightBorderLineStyle = borderLineStyle;
			}
			for (int i = FirstColumn; i <= LastColumn; i++)
			{
				sheet[FirstRow, i].Formatting.TopBorderLineColor = borderColor;
				sheet[FirstRow, i].Formatting.TopBorderLineStyle = borderLineStyle;
				sheet[LastRow, i].Formatting.BottomBorderLineColor = borderColor;
				sheet[LastRow, i].Formatting.BottomBorderLineStyle = borderLineStyle;
			}
			return this;
		}

		/// <summary>
		/// Draw table.
		/// </summary>
		/// <param name="borderColor">Border color.</param>
		/// <param name="lineStyle">Line style.</param>
		/// <returns>Current range instance.</returns>
		public CellRange DrawTable(Color borderColor, BorderLineStyle lineStyle)
		{
			return this.DrawTable(borderColor, lineStyle, lineStyle);
		}

		/// <summary>
		/// Draw table.
		/// </summary>
		/// <param name="borderColor">Border color.</param>
		/// <param name="innerLineStyle">Inner line style.</param>
		/// <param name="outerLineStyle">Outer line style.</param>
		/// <returns>Current range instance.</returns>
		public CellRange DrawTable(Color borderColor, BorderLineStyle innerLineStyle, BorderLineStyle outerLineStyle)
		{
			for (int r = FirstRow; r <= LastRow; r++)
			{
				for (int c = FirstColumn; c <= LastColumn; c++)
				{
					if ( r == FirstRow )
					{
						sheet[r, c].Formatting.TopBorderLineColor = borderColor;
						sheet[r, c].Formatting.TopBorderLineStyle = outerLineStyle;
					}

					if ( c == FirstColumn )
					{
						sheet[r, c].Formatting.LeftBorderLineColor = borderColor;
						sheet[r, c].Formatting.LeftBorderLineStyle = outerLineStyle;
					}

					sheet[r, c].Formatting.RightBorderLineColor = borderColor;
					sheet[r, c].Formatting.RightBorderLineStyle = (c == LastColumn ? outerLineStyle : innerLineStyle);

					sheet[r, c].Formatting.BottomBorderLineColor = borderColor;
					sheet[r, c].Formatting.BottomBorderLineStyle = (r == LastRow ? outerLineStyle : innerLineStyle);
				}
			}
			return this;
		}

		/// <summary>
		/// Set solid background color.
		/// </summary>
		/// <param name="color">Color.</param>
		/// <returns>Current range instance.</returns>
		public CellRange SetBackground(Color color)
		{
			return SetBackground(color, CellBackgroundPattern.Fill);
		}

		/// <summary>
		/// Set backround color and pattern.
		/// </summary>
		/// <param name="color">Color.</param>
		/// <param name="pattern">Pattern.</param>
		/// <returns>Current range instance.</returns>
		public CellRange SetBackground(Color color, CellBackgroundPattern pattern)
		{
			for (int r = FirstRow; r <= LastRow; r++)
			{
				for (int c = FirstColumn; c <= LastColumn; c++)
				{
					sheet[r, c].Formatting.BackgroundPatternColor = color;
					sheet[r, c].Formatting.BackgroundPattern = pattern;
				}
			}
			return this;
		}

		/// <summary>
		/// Set cell alignment.
		/// </summary>
		/// <param name="horizontal">Horizontal alignment.</param>
		/// <param name="vertical">Verical alignment.</param>
		/// <returns>Current range instance.</returns>
		public CellRange SetAlignment(CellHorizontalAlignment horizontal, CellVerticalAlignment vertical)
		{
			for (int r = FirstRow; r <= LastRow; r++)
			{
				for (int c = FirstColumn; c <= LastColumn; c++)
				{
					sheet[r, c].Formatting.HorizontalAlignment = horizontal;
					sheet[r, c].Formatting.VerticalAlignment = vertical;
				}
			}
			return this;
		}

        /// <summary>
        /// Set font.
        /// </summary>
        /// <param name="font">Font.</param>
        /// <returns>Current range instance.</returns>
        public CellRange SetFont(Font font)
        {
            for (int r = FirstRow; r <= LastRow; r++)
            {
                for (int c = FirstColumn; c <= LastColumn; c++)
                {
                    sheet[r, c].Formatting.Font.CopyValuesFrom(font);
                }
            }
            return this;
        }

        /// <summary>
        /// Set cell style.
        /// </summary>
        /// <param name="font">Font.</param>
        /// <returns>Current range instance.</returns>
        public CellRange SetFormatting(Style style)
        {
            for (int r = FirstRow; r <= LastRow; r++)
            {
                for (int c = FirstColumn; c <= LastColumn; c++)
                {
                    sheet[r, c].Formatting.CopyValuesFrom(style);
                }
            }
            return this;
        }
    }
}
