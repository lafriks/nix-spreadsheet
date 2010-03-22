using System;
using System.Collections;
using System.Data;
using System.Text;
using System.Security.Cryptography;

using Nix.SpreadSheet;
using Nix.Zip;

namespace Test
{
    /// <summary>
    /// Summary description for Test.
    /// </summary>
    public class Test
    {
        // Generic function using HashAlgorithm 
        public static void HashData(HashAlgorithm hashAlg, string str )
        {
            byte[] rawBytes = System.Text.ASCIIEncoding.ASCII.GetBytes( str );
            byte[] hashData = hashAlg.ComputeHash( rawBytes );
            Console.WriteLine( BitConverter.ToString( hashData ) );
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            /*Adler32 ad32 = new Adler32();
            HashData(ad32, "Wikipedia");
            CRC32 crc32 = new CRC32(CRC32.DefaultPolynomialLE);
            HashData(crc32, "Wikipedia");*/
            SpreadSheetDocument doc = new SpreadSheetDocument();
            Sheet s = doc.AddSheet();
            s["A1"].Formatting.WrapText = true;
            s["A1"].Value = 13;
            s["A1"].Formatting.BackgroundPatternColor = System.Drawing.Color.Red;
            s["A1"].Formatting.BackgroundPattern = CellBackgroundPattern.Fill;
            s["B2"].Value = "Test";
            s["B3"].Value = "TestX";
            s["B4"].Value = "TestXX";
            s["B5"].Value = "Test";
            s["B6"].Value = "TestXX";
            s.GetCellRange("B2:E6").DrawTable(System.Drawing.Color.Black, BorderLineStyle.Thin, BorderLineStyle.Medium);
            s.GetCellRange("B2:B6").SetBackground(System.Drawing.Color.Gray)
            					   .DrawBorder(System.Drawing.Color.Black, BorderLineStyle.Medium)
            					   .SetAlignment(CellHorizontalAlignment.Centred, CellVerticalAlignment.Centred);
            s["B2"].Formatting.LeftBorderLineColor = System.Drawing.Color.Black;
            s.Columns[0].Width = 256 * 10;
            s.Columns[3].Formatting.BackgroundPattern = CellBackgroundPattern.Fill;
            s.Columns[3].Formatting.BackgroundPatternColor = System.Drawing.Color.Gray;
            //s["IV65536"].Value = "MAX";
			Sheet big = doc.AddSheet("Big");
			big.InsertTable(0, 0, BigTestData().DefaultView);
            doc.Save(@"test.xls", new Nix.SpreadSheet.Provider.XlsFileFormatProvider());
        }

		public static DataTable BigTestData()
		{
			Random rnd = new Random();
			DataTable dt = new DataTable();
			dt.Columns.Add("nr", typeof(int));
			dt.Columns.Add("text1", typeof(string));
			dt.Columns.Add("text2", typeof(string));
			dt.Columns.Add("text3", typeof(string));
			dt.Columns.Add("text4", typeof(string));
			dt.Columns.Add("text5", typeof(string));
			dt.Columns.Add("text6", typeof(string));
			dt.Columns.Add("text7", typeof(string));
			dt.Columns.Add("text8", typeof(string));
			dt.Columns.Add("text9", typeof(string));
			dt.Columns.Add("text10", typeof(string));
			dt.Columns.Add("text11", typeof(string));

			for (int i = 0; i < 4000; i++)
			{
				dt.Rows.Add(new object[] { i,
									"Col1_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col2_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col3_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col4_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col5_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col6_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col7_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col8_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col9_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col10_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString(),
									"Col11_" + ((int)Math.Round(rnd.NextDouble() * 10, 0)).ToString()
							});
			}

			return dt;
		}
	}
}
