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

namespace Nix.SpreadSheet.Common
{
	/// <summary>
	/// Description of Position.
	/// </summary>
	public struct Position : IEquatable<Position>
	{
		public Position(int row, int column)
			: this(false)
		{
			this.row = row;
			this.column = column;
		}

		private bool empty;
		
		private Position(bool empty)
		{
			this.empty = empty;
			this.row = -1;
			this.column = -1;
		}

		public static readonly Position Empty;

		static Position()
		{
			Empty = new Position(false);
		}

		public bool IsEmpty()
		{
			return this.empty;
		}

		private int row;
		
		public int Row {
			get { return row; }
			set { row = value; }
		}
		
		private int column;
		
		public int Column {
			get { return column; }
			set { column = value; }
		}

		#region Equals and GetHashCode implementation
		public override bool Equals(object obj)
		{
			if (obj is Position)
				return Equals((Position)obj); // use Equals method below
			else
				return false;
		}
		
		public bool Equals(Position other)
		{
			return this.Row == other.Row && this.Column == other.Column;
		}

		public override int GetHashCode()
		{
			return this.Column.GetHashCode() ^ this.Row.GetHashCode();
		}
		
		public static bool operator ==(Position left, Position right)
		{
			return left.Equals(right);
		}
		
		public static bool operator !=(Position left, Position right)
		{
			return !left.Equals(right);
		}
		#endregion
	}
}
