using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Spreadsheet document.
	/// </summary>
	public class SpreadSheetDocument : IEnumerable
	{
		#region Sheets
		private Dictionary<string, Sheet> sheets = new Dictionary<string,Sheet>();

		public Sheet AddSheet ()
		{
			string name = string.Empty;
			for ( int i = 1; ; i++ )
			{
				name = "Sheet" + i.ToString();
				if ( !this.sheets.ContainsKey(name) )
					break;
			}
			this.sheets.Add(name, new Sheet(this, name));
			return this.sheets[name];
		}

		public Sheet AddSheet (string name)
		{
			if (this.sheets.ContainsKey(name))
				throw new ArgumentException("Sheet with such name already exists.");
			this.sheets.Add(name, new Sheet(this, name));
			return this.sheets[name];
		}

		internal void ChangeSheetName ( string oldName, string newName )
		{
			if (this.sheets.ContainsKey(newName))
				throw new ArgumentException("Sheet with such name already exists.");
			Sheet s = this.sheets[oldName];
			this.sheets.Remove(oldName);
			this.sheets.Add(newName, s);
		}

		/// <summary>
		/// Gets the sheet count.
		/// </summary>
		/// <value>The sheet count.</value>
		public int SheetCount
		{
			get { return this.sheets.Count; }
		}
		#endregion

		#region Loading and saving spreadsheet documents
		private static Provider.IFileFormatProvider GetFileFormatProvider(SpreadSheetFileFormat format)
		{
			switch(format)
			{
				case SpreadSheetFileFormat.ExcelBinary:
					return new Provider.Xls();
				default:
					throw new NotImplementedException("Unsupported spreadsheet file format.");
			}
		}
		
		/// <summary>
		/// Loads a spreadsheet document from the specified file.
		/// </summary>
		/// <param name="fileName">The file name.</param>
		/// <param name="format">The format.</param>
		/// <returns>Spreadsheet document.</returns>
		public static SpreadSheetDocument Load ( string fileName, SpreadSheetFileFormat format )
		{
			StreamReader s = new StreamReader(fileName);
			try
			{
				return Load(s.BaseStream, format);
			}
			finally
			{
				s.Close();
			}
		}

		/// <summary>
		/// Loads a spreadsheet document from the specified stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		/// <param name="format">The format.</param>
		/// <returns>Spreadsheet document.</returns>
		public static SpreadSheetDocument Load ( Stream stream, SpreadSheetFileFormat format )
		{
			return new SpreadSheetDocument();
		}

		/// <summary>
		/// Saves spreadsheet document to the specified file.
		/// </summary>
		/// <param name="fileName">The file name.</param>
		/// <param name="format">The format.</param>
		public void Save ( string fileName, SpreadSheetFileFormat format )
		{
			StreamWriter s = new StreamWriter(fileName);
			try
			{
				this.Save(s.BaseStream, format);
			}
			finally
			{
				s.Close();
			}
		}

		/// <summary>
		/// Saves the spreadsheet document to the stream.
		/// </summary>
		/// <param name="stream">The stream.</param>
		/// <param name="format">The format.</param>
		public void Save ( Stream stream, SpreadSheetFileFormat format )
		{
			Provider.IFileFormatProvider prov = GetFileFormatProvider(format);
			prov.Save(this, stream);
		}
		#endregion

		#region IEnumerable Members

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator ()
		{
			return this.sheets.Keys.GetEnumerator();
		}

		#endregion
	}
}
