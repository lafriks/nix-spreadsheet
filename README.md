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

## Tests

The library retains its .NET Standard 2.0 target and supports spreadsheet export
and automatic row heights on Windows and Linux, including Alpine, on .NET 6, 8,
and 10, and on Windows with .NET Framework 4.8. All platforms measure text with
the managed SixLabors.Fonts library, using the requested installed font or an
embedded Liberation Sans fallback when
that font is unavailable. No native graphics packages or system fonts need to be
installed in Linux containers.
Font availability can produce different measured heights across platforms.

The SixLabors.Fonts dependency is restricted to `[1.0.0,2.0.0)` to exclude
releases under the Six Labors Split License. Version 1.0.0 uses the
[Apache 2.0 license](https://github.com/SixLabors/Fonts/blob/v1.0.0/LICENSE).

Version 0.4.0 changes `Font.ToNativeFont()` to return `SixLabors.Fonts.Font`
instead of `System.Drawing.Font`. Existing callers must rebuild and migrate any
GDI+ rendering code. The unused `Font.ToNativeWinFont()` conversion helper and
the System.Drawing.Common dependency have been removed. The portable font carries
the family, point size, bold and italic style; underline, strikeout and charset remain spreadsheet formatting
properties and are still written to the workbook. Automatic row heights use
this portable conversion helper on every platform.
The fallback fonts and their redistribution license come from
[Liberation Fonts 2.1.5](https://github.com/liberationfonts/liberation-fonts/releases/tag/2.1.5).

The regression test project targets .NET 8 by default. Install a .NET 8 SDK or
newer and a .NET 8 runtime, then run:

```bash
cd SpreadSheet.Tests
dotnet test
```

The tests generate `.xls` files with synthetic data and read them with NPOI to
verify their contents and formatting. They require no database or Microsoft Excel.
All tests run on every platform, including automatic row height measurements.

The [GitHub Actions workflow](.github/workflows/tests.yml) runs the complete suite
on Windows, Ubuntu, and Linux Alpine with .NET 6, 8, and 10, plus .NET Framework
4.8 on Windows, for pushes to `master` and pull requests. It can also be started
manually. Its shared script selects the test framework through command-line
properties, leaving the test project
file unchanged and keeping the library references on .NET Standard 2.0:

```bash
bash .github/scripts/test.sh 10.0
```

For legacy validation, run `bash .github/scripts/test.sh 6.0`, or on Windows with
the .NET Framework 4.8 targeting pack, `bash .github/scripts/test.sh net48`.
The Framework job uses Windows Server 2022 and logs its installed runtime release.

See [the test coverage notes](SpreadSheet.Tests/README.md) for the covered scenarios.
