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

## Deployment on Linux

All platforms measure text with SkiaSharp and HarfBuzz. Font lookup uses the
system font manager, including Fontconfig aliases on Linux. An embedded
Liberation Sans font supplies regular, bold, italic and bold italic styles when
the requested family cannot be resolved. Font availability can produce different
measured heights across platforms. SkiaSharp uses the
[MIT license](https://github.com/mono/SkiaSharp/blob/main/LICENSE.md).

Linux applications must reference `SkiaSharp.NativeAssets.Linux` 4.153.1 and
`HarfBuzzSharp.NativeAssets.Linux` 14.2.1.301 in the executable project, and install
Fontconfig in the runtime image. Use `linux-musl-x64` or `linux-musl-arm64` when
publishing for Alpine. The `SkiaSharp.NativeAssets.Linux.NoDependencies` package
does not support Fontconfig substitution.

For Alpine, install the font dependencies with:

```bash
apk add --no-cache fontconfig ttf-liberation font-carlito
fc-cache -f
```

These packages provide Liberation Sans for Arial, Liberation Serif for Times New
Roman, and Carlito for Calibri. The example Dockerfiles install these substitutes.
The workbook retains the originally requested font names.

`Font.ToNativeFont()` returns an `SKFont` with its size in points; callers must
dispose it. Automatic row heights convert this size to pixels at 96 DPI.
Underline, strikeout and charset remain spreadsheet formatting properties and
are written to the workbook.
The fallback fonts and their redistribution license come from
[Liberation Fonts 2.1.5](https://github.com/liberationfonts/liberation-fonts/releases/tag/2.1.5).

## Tests

The library retains its .NET Standard 2.0 target and supports spreadsheet export
and automatic row heights on Windows and Linux, including Alpine, on .NET 6, 8,
and 10, and on Windows with .NET Framework 4.8.

The regression test project targets .NET 8 by default. Install a .NET 8 SDK or
newer and a .NET 8 runtime, then run:

```bash
cd SpreadSheet.Tests
dotnet test
```

The tests generate `.xls` files with synthetic data and read them with NPOI to
verify their contents and formatting. They require no database or Microsoft Excel.
Cross-platform tests, including automatic row height measurements, run on all
configured platforms. Tests requiring Windows-installed fonts run only on
Windows.

The [GitHub Actions workflow](.github/workflows/tests.yml) runs the tests on
Windows, Ubuntu, and Linux Alpine with .NET 6, 8, and 10, plus .NET Framework 4.8
on Windows, for pushes to `master` and pull requests. Tests requiring
Windows-installed fonts are skipped on other platforms. The workflow can also be
started manually. Its shared script selects the test framework through command-line
properties, leaving the test project file unchanged and keeping the library
references on .NET Standard 2.0:

```bash
bash .github/scripts/test.sh 10.0
```

For legacy validation, run `bash .github/scripts/test.sh 6.0`, or on Windows with
the .NET Framework 4.8 targeting pack, `bash .github/scripts/test.sh net48`.

See [the test coverage notes](SpreadSheet.Tests/README.md) for the covered scenarios.
