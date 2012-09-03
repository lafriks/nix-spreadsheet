using System;
using System.Collections;
using System.Data;
using System.Text;
using System.Security.Cryptography;

using Nix.SpreadSheet;
using System.Drawing;

namespace Test
{
    /// <summary>
    /// Summary description for Test.
    /// </summary>
    public class Test
    {
        // Generic function using HashAlgorithm 
        /*public static void HashData(HashAlgorithm hashAlg, string str )
        {
            byte[] rawBytes = System.Text.ASCIIEncoding.ASCII.GetBytes( str );
            hashAlg.ComputeHash( rawBytes );
			Console.WriteLine(BitConverter.ToString(hashAlg.Hash));
        }*/

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
            s.PageSettings.Landscape = false;
            s.PageSettings.PapeSize = Nix.SpreadSheet.PageSettings.PageSize.A5;
            s["A1"].Formatting.WrapText = true;
            s["A1"].Formatting.BackgroundPattern = CellBackgroundPattern.Fill;
            s["A1"].Formatting.BackgroundPatternColor = Color.Pink;
            s.GetCellRange("A1:A1").DrawBorder(Color.Red, BorderLineStyle.Thin);
            s["A1"].Value = (decimal)13.02;
            s["A1"].Formatting.Format = "+0.00;[Red]-0.00;0";
            //s["A1"].Formatting.BackgroundPatternColor = System.Drawing.Color.Red;
            //s["A1"].Formatting.BackgroundPattern = CellBackgroundPattern.Fill;
            s["B2"].Value = (decimal)-7012.12;
            s["B2"].Formatting.Format = "+0.00;[Red]-0.00;0";
            s["B3"].Value = "TestX";
            s["B4"].Value = "TestXX sdsd sd sd sd sdsd sdsd sdsddd" + Environment.NewLine + "sdfsdf sddf gdffdgdfgfdgdfgdfgfdgfdgg dfgdfg gdfgdfgf sdfsdfdsf sdf";
            s["B4"].Formatting.WrapTextAtRightBorder = true;
            s.GetCellRange("D5:H7").Merge();
            s["D5"].Value = "TestXX sdsd sd sd sd sdsd sdsd sdsddd" + Environment.NewLine + "sdfsdf sddf gdffdgdfgfdgdfgdfgfdgfdgg dfgdfg gdfgdfgf sdfsdfdsf sdf" + Environment.NewLine + "sdfsdf sddf gdffdgdfgfdgdfgdfgfdgfdgg dfgdfg gdfgdfgf sdfsdfdsf sdf";
            s["D5"].Formatting.WrapTextAtRightBorder = true;
            s["D5"].Formatting.VerticalAlignment = CellVerticalAlignment.Top;
            s.GetCellRange("D5:H7").DrawBorder(Color.Black, BorderLineStyle.Medium);
            s["B5"].Value = "Test";
            s["B6"].Value = "TestXX";

            s.AutoSizeRowHeight();
            /*s.GetCellRange("B2:E6").DrawTable(System.Drawing.Color.Black, BorderLineStyle.Thin, BorderLineStyle.Medium);
            s.GetCellRange("B2:B6").SetBackground(System.Drawing.Color.Gray)
            					   .DrawBorder(System.Drawing.Color.Black, BorderLineStyle.Medium)
            					   .SetAlignment(CellHorizontalAlignment.Centred, CellVerticalAlignment.Centred);
            s["B2"].Formatting.LeftBorderLineColor = System.Drawing.Color.Black;
            s.Columns[0].Width = 1000;
            s.Columns[1].Width = 913;
            s[1].Height = 20;
            s[10].Height = 40;
            s.Columns[3].Formatting.BackgroundPattern = CellBackgroundPattern.Fill;
            s.Columns[3].Formatting.BackgroundPatternColor = System.Drawing.Color.Gray;
            //s["IV65536"].Value = "MAX";
			Sheet big = doc.AddSheet("Big");
			big.InsertTable(0, 0, BigTestData().DefaultView);
            big.GetCellRange(0, 0, 1, 3).Merge();*/
            doc.Save(@"test.xls", new Nix.SpreadSheet.Provider.XlsFileFormatProvider());
			//doc.Save(@"test.ods", new Nix.SpreadSheet.Provider.OpenDocumentFileFormatProvider());
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

			for (int i = 0; i < 30000; i++)
			{
				dt.Rows.Add(new object[] { i,
									"Col_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Pīp_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Šūņš_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Šūns_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Ban_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Rīb_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Bak8_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Rot9_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Žog10_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Māk11_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString(),
									"Sāk12_" + ((int)Math.Round(rnd.NextDouble() * 10000, 0)).ToString()
							});
			}

			return dt;
		}
	}
}
