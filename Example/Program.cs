using Nix.SpreadSheet;
using System;
using System.Data;
using System.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
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
            s.GetCellRange("D5:H5").Merge();
            s["D5"].Value = "TestXX sdsd sd sd sd sdsd sdsd sdsddd" + Environment.NewLine + "sdfsdf sddf gdffdgdfgfdgdfgdfgfdgfdgg dfgdfg gdfgdfgf sdfsdfdsf sdf" + Environment.NewLine + "sdfsdf sddf gdffdgdfgfdgdfgdfgfdgfdgg dfgdfg gdfgdfgf sdfsdfdsf sdf";
            s["D5"].Formatting.WrapTextAtRightBorder = true;
            s["D5"].Formatting.VerticalAlignment = CellVerticalAlignment.Top;
            s.GetCellRange("D5:H5").DrawBorder(Color.Black, BorderLineStyle.Medium);
            s["B5"].Value = "Test";
            s["B6"].Value = "TestXX";

            s.Columns[5].Width = 200;
            s.GetCellRange("A8:B8").Merge();
            s["C8"].Value = "Jānis Testiņš";
            s["D8"].Value = "Test";
            s["E8"].Value = "Pētris Testiņš";
            s["F8"].Value = "fffg dddddddddddddddd aaaaaaaaaaaaaaaaaa dddddddddd";
            s["F8"].Formatting.WrapTextAtRightBorder = true;
            s["G8"].Value = new DateTime(2012, 9, 3);
            s["G8"].Formatting.Format = "dd.MM.yyyy";

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

            CreateFontSizeExample();
            CreateAutoSizeRowHeightExample();
        }

        private static void CreateFontSizeExample()
        {
            SpreadSheetDocument doc = new SpreadSheetDocument();
            Sheet sheet = doc.AddSheet("Font sizes");
            sheet.Columns[0].Width = 150;
            sheet.Columns[1].Width = 100;
            sheet.Columns[2].Width = 300;
            sheet[0, 0].Value = "Font";
            sheet[0, 1].Value = "Size (pt)";
            sheet[0, 2].Value = "Sample text";

            Sheet longTextSheet = doc.AddSheet("Long wrapped text");
            longTextSheet.Columns[0].Width = 150;
            longTextSheet.Columns[1].Width = 100;
            longTextSheet.Columns[2].Width = 300;
            longTextSheet[0, 0].Value = "Font";
            longTextSheet[0, 1].Value = "Size (pt)";
            longTextSheet[0, 2].Value = "Very long wrapped text";
            string longText = "Šis ir ļoti garš teksta piemērs automātiskas rindu augstuma aprēķināšanas pārbaudei. "
                + "Vienā šūnā ir vairāki teikumi ar latviešu burtiem: āčēģīķļņšūž. "
                + "Tekstam jāaplaužas pie kolonnas labās malas un pilnībā jāietilpst rindā neatkarīgi no fonta un tā izmēra. "
                + "Šeit ir arī ļoti garš vārds bez atstarpēm: ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789. "
                + "Pēdējam teikumam jābūt redzamam bez manuālas rindas augstuma pielāgošanas.";

            string[] fontNames = { "Times New Roman", "Calibri" };
            ushort[] fontSizes = { 100, 200, 280, 360 };
            int row = 1;
            foreach (string fontName in fontNames)
            {
                foreach (ushort fontSize in fontSizes)
                {
                    sheet[row, 0].Value = fontName;
                    sheet[row, 1].Value = fontSize / 20;
                    sheet[row, 2].Value = "Latviešu teksts: āčēģīķļņšūž\nOtrā rinda ar lielāku fontu";
                    Nix.SpreadSheet.Font font = new Nix.SpreadSheet.Font
                    {
                        Name = fontName,
                        Size = fontSize
                    };
                    sheet.GetCellRange(row, 0, row, 2).SetFont(font);
                    sheet.GetCellRange(row, 0, row, 2).DrawBorder(Color.Gray, BorderLineStyle.Thin);
                    sheet[row, 2].Formatting.WrapTextAtRightBorder = true;

                    longTextSheet[row, 0].Value = fontName;
                    longTextSheet[row, 1].Value = fontSize / 20;
                    longTextSheet[row, 2].Value = longText;
                    longTextSheet.GetCellRange(row, 0, row, 2).SetFont(font);
                    longTextSheet.GetCellRange(row, 0, row, 2).DrawBorder(Color.Gray, BorderLineStyle.Thin);
                    longTextSheet[row, 2].Formatting.WrapTextAtRightBorder = true;
                    row++;
                }
            }

            sheet[1].Height = 48;
            sheet.AutoSizeRowHeight();
            longTextSheet.AutoSizeRowHeight();
            doc.Save(@"test-font-sizes.xls", new Nix.SpreadSheet.Provider.XlsFileFormatProvider());
        }

        private static void CreateAutoSizeRowHeightExample()
        {
            SpreadSheetDocument doc = new SpreadSheetDocument();
            Sheet sheet = doc.AddSheet("Row heights");
            sheet.Columns[0].Width = 180;
            sheet.Columns[1].Width = 300;

            sheet[0, 0].Value = "Times New Roman, 10 pt";
            sheet[0, 1].Value = "This row keeps its explicitly assigned height.\nIt has more text than the default row height needs.";
            sheet.GetCellRange("A1:B1").SetFont(new Nix.SpreadSheet.Font { Name = "Times New Roman", Size = 200 });
            sheet.GetCellRange("A1:B1").DrawBorder(Color.Gray, BorderLineStyle.Thin);
            sheet[0, 1].Formatting.WrapTextAtRightBorder = true;
            sheet[0].Height = 60;

            sheet[1, 0].Value = "Calibri, 14 pt";
            sheet[1, 1].Value = "Automatic height fits this multiline text.\nThe font is applied to the whole cell range.";
            sheet.GetCellRange("A2:B2").SetFont(new Nix.SpreadSheet.Font { Name = "Calibri", Size = 280 });
            sheet.GetCellRange("A2:B2").DrawBorder(Color.Gray, BorderLineStyle.Thin);
            sheet[1, 1].Formatting.WrapTextAtRightBorder = true;

            sheet[2, 0].Value = "Times New Roman, 18 pt";
            sheet[2, 1].Value = "A larger font increases the automatic row height.\nExplicit heights on other rows are preserved.";
            sheet.GetCellRange("A3:B3").SetFont(new Nix.SpreadSheet.Font { Name = "Times New Roman", Size = 360 });
            sheet.GetCellRange("A3:B3").DrawBorder(Color.Gray, BorderLineStyle.Thin);
            sheet[2, 1].Formatting.WrapTextAtRightBorder = true;

            sheet[3, 0].Value = "Explicit height";
            sheet[3, 1].Value = "This row is also left at its specified height.";
            sheet.GetCellRange("A4:B4").SetFont(new Nix.SpreadSheet.Font { Name = "Calibri", Size = 240 });
            sheet.GetCellRange("A4:B4").DrawBorder(Color.Gray, BorderLineStyle.Thin);
            sheet[3].Height = 36;

            sheet.AutoSizeRowHeight();
            doc.Save(@"test-auto-size-row-height.xls", new Nix.SpreadSheet.Provider.XlsFileFormatProvider());
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
