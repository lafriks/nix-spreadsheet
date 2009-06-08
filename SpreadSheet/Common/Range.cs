/*
 * Created by SharpDevelop.
 * User: User
 * Date: 2009.04.05.
 * Time: 17:50
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace Nix.SpreadSheet.Common
{
	public struct Range
	{
		public Range(Position p_Start, Position p_End):this(p_Start, p_End, true)
		{
		}

		public Range(int p_StartRow, int p_StartCol, int p_EndRow, int p_EndCol)
		{
			m_Start = new Position(p_StartRow, p_StartCol);
			m_End = new Position(p_EndRow, p_EndCol);
			Normalize();
		}

		private Range(Position p_Start, Position p_End, bool p_bCheck)
		{
			m_Start = p_Start;
			m_End = p_End;

			if (p_bCheck)
				Normalize();
		}

		static Range()
		{
			Empty = new Range(Position.Empty, Position.Empty, false);
		}

		private Position m_Start, m_End;

		public Position Start
		{
			get{return m_Start;}
		}

		public Position End
		{
			get{return m_End;}
		}

		public void MoveTo(Position p_StartPosition)
		{
			int l_ColCount = ColumnsCount;
			int l_RowCount = RowsCount;
			m_Start = p_StartPosition;
			RowsCount = l_RowCount;
			ColumnsCount = l_ColCount;
		}

		public int ColumnsCount
		{
			get
			{
				return (m_End.Column - m_Start.Column)+1;
			}
			set
			{
				if (value<=0)
					throw new ArgumentOutOfRangeException("ColumnsCount");
				m_End = new Position(m_End.Row, m_Start.Column+value-1);
			}
		}

		public int RowsCount
		{
			get
			{
				return (m_End.Row - m_Start.Row)+1;
			}
			set
			{
				if (value<=0)
					throw new ArgumentOutOfRangeException("RowsCount");
				m_End = new Position(m_Start.Row+value-1, m_End.Column);
			}
		}

		public Range(Position p_SinglePosition):this(p_SinglePosition, p_SinglePosition, false)
		{
		}

		public readonly static Range Empty;

		private void Normalize()
		{
			int l_MinRow, l_MinCol, l_MaxRow, l_MaxCol;
			
			if (m_Start.Row < m_End.Row)
				l_MinRow = m_Start.Row;
			else
				l_MinRow = m_End.Row;

			if (m_Start.Column < m_End.Column)
				l_MinCol = m_Start.Column;
			else
				l_MinCol = m_End.Column;


			if (m_Start.Row > m_End.Row)
				l_MaxRow = m_Start.Row;
			else
				l_MaxRow = m_End.Row;

			if (m_Start.Column > m_End.Column)
				l_MaxCol = m_Start.Column;
			else
				l_MaxCol = m_End.Column;

			m_Start = new Position(l_MinRow, l_MinCol);
			m_End = new Position(l_MaxRow, l_MaxCol);
		}

		public bool ContainsRow(int p_Row)
		{
			return (p_Row >= m_Start.Row &&
				p_Row <= m_End.Row);
		}

		public bool ContainsColumn(int p_Col)
		{
			return (p_Col >= m_Start.Column &&
				p_Col <= m_End.Column);
		}

		public bool Contains(Position p_Position)
		{
			int l_Row = p_Position.Row;
			int l_Col = p_Position.Column;

			return (l_Row >= m_Start.Row &&
					l_Col >= m_Start.Column &&
					l_Row <= m_End.Row &&
					l_Col <= m_End.Column);
		}

		public bool Contains(Range p_Range)
		{
			return (Contains(p_Range.Start) &&
					Contains(p_Range.End));
		}

		public bool IsEmpty()
		{
			return (Start.IsEmpty() || End.IsEmpty());
		}

		public override int GetHashCode()
		{
			return this.Start.GetHashCode() ^ this.End.GetHashCode();
		}

		public static bool operator == (Range Left, Range Right)
		{
			return Left.Equals(Right);
		}

		public static bool operator != (Range Left, Range Right)
		{
			return !Left.Equals(Right);
		}

		public bool Equals(Range p_Range)
		{
			return (Start.Equals(p_Range.Start) && End.Equals(p_Range.End));
		}

		public override bool Equals(object obj)
		{
			return Equals((Range)obj);
		}

		public override string ToString()
		{
			return Start.ToString() + " to " + End.ToString();
		}
	}
}
