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
| Row heights | Multiline and merged titles fit, explicitly assigned heights survive the default autosize call, heights survive serialization |

The tests use fixed dates and an explicit Latvian document locale. They require
no database, production data, spreadsheet fixtures or Excel installation.

The project targets `netcoreapp3.1` and uses C# 8 syntax. Its local `global.json`
selects SDK `3.1.426`, the oldest SDK installed on the development device, and
allows later 3.1 feature bands without rolling forward to a newer major version.
The library projects retain their `netstandard2.0` targets.

Run from the repository root:

```bash
cd SpreadSheet.Tests
dotnet --version
dotnet test
```


Only the row height test needs Windows, due to the library's existing
System.Drawing.Common dependency. It is explicitly skipped on other platforms.
No exact font measurement is asserted: the test compares multiline heights with
single line heights to allow for differences in installed fonts and DPI.

When extending coverage, construct workbooks with synthetic data and add
assertions on the contents and formatting read by NPOI.
