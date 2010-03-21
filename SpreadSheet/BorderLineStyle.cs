/*
 * Library for generating spreadsheet documents.
 * Copyright (C) 2008, Lauris Bukðis-Haberkorns <lauris@nix.lv>
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
	/// Border line style.
	/// </summary>
	public enum BorderLineStyle
	{
		/// <summary>
		/// None.
		/// </summary>
		None = 0,
		/// <summary>
		/// Thin.
		/// </summary>
		Thin,
		/// <summary>
		/// Medium.
		/// </summary>
		Medium,
		/// <summary>
		/// Dashed.
		/// </summary>
		Dashed,
		/// <summary>
		/// Dotted.
		/// </summary>
		Dotted,
		/// <summary>
		/// Thick.
		/// </summary>
		Thick,
		/// <summary>
		/// Double.
		/// </summary>
		Double,
		/// <summary>
		/// Hair.
		/// </summary>
		Hair,
		/// <summary>
		/// Medium dashed.
		/// </summary>
		MediumDashed,
		/// <summary>
		/// Thin dash dotted.
		/// </summary>
		ThinDashDotted,
		/// <summary>
		/// Medium dash dotted.
		/// </summary>
		MediumDashDotted,
		/// <summary>
		/// Thin dash dot dotted.
		/// </summary>
		ThinDashDotDotted,
		/// <summary>
		/// Medium dash dot dotted.
		/// </summary>
		MediumDashDotDotted,
		/// <summary>
		/// Slanted medium dash dotted
		/// </summary>
		SlantedMediumDashDotted
	}
}
