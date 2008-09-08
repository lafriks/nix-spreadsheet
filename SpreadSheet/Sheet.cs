using System;
using System.Collections.Generic;
using System.Text;

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Sheet.
	/// </summary>
	public class Sheet
	{
		private SpreadSheetDocument document;

		/// <summary>
		/// Gets the owner document.
		/// </summary>
		/// <value>The owner document.</value>
		public SpreadSheetDocument Document
		{
			get { return document; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Sheet"/> class.
		/// </summary>
		/// <param name="document">The owner document.</param>
		/// <param name="name">The name.</param>
		public Sheet(SpreadSheetDocument document, string name)
		{
			this.document = document;
			this.name = name;
		}

		private string name = "Sheet";

		/// <summary>
		/// Gets or sets the sheet name.
		/// </summary>
		/// <value>The sheet name.</value>
		public string Name
		{
			get { return name; }
			set
			{
				this.Document.ChangeSheetName(name, value);
				name = value;
			}
		}
	}
}
