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
