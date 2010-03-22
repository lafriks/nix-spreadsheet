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

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Error code.
	/// </summary>
	public enum ErrorCode
	{
		/// <summary>
		/// Intersection of two cell ranges is empty.
		/// </summary>
		Null,
		/// <summary>
		/// Division by zero.
		/// </summary>
		DivisionByZero,
		/// <summary>
		/// Wrong type of operand.
		/// </summary>
		WrongType,
		/// <summary>
		/// Illegal or deleted cell reference.
		/// </summary>
		IllegalCellReference,
		/// <summary>
		/// Wrong function or range name.
		/// </summary>
		WrongFunctionOrRange,
		/// <summary>
		/// Value range overflow.
		/// </summary>
		ValueRangeOverflow,
		/// <summary>
		/// Argument or function not available.
		/// </summary>
		ArgumentOrFunctionNotAvailable
	}
}
