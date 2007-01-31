using System;
using System.Collections;
using System.Data;
using System.Text;

using Nix.SpreadSheet;

namespace Test
{
	/// <summary>
	/// Summary description for Test.
	/// </summary>
	public class Test
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
        	SpreadSheet sp = new SpreadSheet();
        	sp.Save("test.xls");
        }
}
