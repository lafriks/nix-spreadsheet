Nix.SpreadSheet
===============

Nix.SpreadSheet is LGPL'ed library for generating files in Microsoft Excel binary format. It's entirely written C# and can be compiled for using in .NET Standard 2.0 projects. It also includes Nix.CompoundFile library for generating files in OLE 2 Compount Document format.

Example code:
```csharp
SpreadSheetDocument doc = new SpreadSheetDocument();

Sheet s = doc.AddSheet();
s["A1"].Formatting.WrapText = true;
s["A1"].Formatting.BackgroundPattern = CellBackgroundPattern.Fill;
s["A1"].Formatting.BackgroundPatternColor = Color.Pink;
s.GetCellRange("A1:A1").DrawBorder(Color.Red, BorderLineStyle.Thin);
s["A1"].Value = (decimal)13.02;
s["A1"].Formatting.Format = "+0.00;[Red]-0.00;0";
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

s.Columns[5].Width = 200;
s.GetCellRange("A8:B8").Merge();
s["G8"].Value = new DateTime(2012, 9, 3);
s["G8"].Formatting.Format = "dd.MM.yyyy";

s.AutoSizeRowHeight();

doc.Save(@"test.xls", new Nix.SpreadSheet.Provider.XlsFileFormatProvider());
```
