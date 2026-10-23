# Excel export regression tests

These tests exercise this repository's library through its public API with
synthetic spreadsheet data.

Each test saves a workbook to a memory stream, rewinds it, and reopens the bytes
with NPOI's independent `.xls` reader. Assertions check the serialized
contents and layout rather than just checking whether saving throws.

| Scenario | Checks |
| --- | --- |
| Document list | Latvian characters, multiline text, registration numbers with leading zeros, decimal amounts, text flags, null and empty cells |
| Registration dates | Dates and timestamps remain strings with the specified format; missing dates remain blank |
| Report table | Numeric counts and totals, fonts, bold headers, wrapping, alignment, inner and outer borders, background palette, pixel column widths |
| Merged headers | Horizontal and vertical merges, anchor values and adjacent detail rows |
| Grouped export | Sheet order, Unicode names, distinct group data, shared text and portrait details sheet |
| Empty groups | Readable fallback worksheet, both empty and with report headers |
| Large document list | All 2,048 rows, unique and repeated Unicode strings, shared string continuation records, multiple OLE sectors, row blocks and a footer after a gap |
| Long document content | A single multiline Unicode string spanning multiple continuation records and the following cell |
| Row heights | Multiline and merged titles fit, explicitly assigned heights survive the default autosize call, heights survive serialization, the multiline report is repeated for several font sizes and weights, row height scales linearly with line count across 10-18 pt regular and bold text for the fallback font on every platform and for Arial, Times New Roman and Calibri on Windows only (skipped elsewhere because the fonts are not installed), and Arial, Times New Roman and Calibri rows match Excel's line pitch |

The tests use fixed dates and an explicit Latvian document locale. They require
no database, production data, spreadsheet fixtures or Excel installation.

The project targets `net8.0` by default. Install a .NET 8 SDK or newer and a
.NET 8 runtime to build and run the tests.
The library projects retain their `netstandard2.0` targets.

Run from the repository root:

```bash
cd SpreadSheet.Tests
dotnet --version
dotnet test
```

The [GitHub Actions workflow](../.github/workflows/tests.yml) runs on pushes to
`master` and on pull requests, and supports manual runs. Its matrix runs the
cross-platform tests on Windows, Ubuntu, and Alpine with .NET 6, 8, and 10, plus
.NET Framework 4.8 on Windows Server only. Tests requiring Windows-installed
fonts run only on Windows. Alpine uses the official SDK container, which includes
ICU and enables globalization so the tests can use their Latvian document locale.
No font or graphics packages are installed.

To select a framework without editing the test project, install its SDK and
run the same script used in CI from the repository root:

```bash
bash .github/scripts/test.sh 10.0
```

Use `6.0` or `8.0` for those .NET versions, or `net48` on Windows with the .NET
Framework 4.8 targeting pack installed. Framework builds select C# 8.0 so the
existing nullable annotations and using declarations compile.

The script overrides the test frameworks during restore, restores the library
references separately with their .NET Standard 2.0 targets, and builds and runs
the tests on the selected framework. The resulting TRX reports are uploaded by CI.

Cross-platform height tests run on every platform. Height calculation uses
SkiaSharp with HarfBuzz and an embedded Liberation Sans fallback for missing
fonts. Linux runs require Fontconfig and the Linux native asset packages;
the CI images install Liberation and Carlito font substitutes. The public font
conversion helper returns an `SKFont` in points, which callers must dispose.
Layout regressions check cell heights for wrapping, sizes and styles.
The layout regressions also exercise missing fonts in regular, bold, italic,
and bold italic styles, font sizes, width-based wrapping, LF and CRLF line
breaks, empty cell minimum heights, and explicit-height override behavior.
Relative layout checks compare multiline heights with single line heights to
allow for differences in installed fonts and DPI.

When extending coverage, construct workbooks with synthetic data and add
assertions on the contents and formatting read by NPOI.
