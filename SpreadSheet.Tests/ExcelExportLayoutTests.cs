using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using Nix.SpreadSheet;
using Xunit;
using static SpreadSheet.Tests.ExcelExportTestHelper;

namespace SpreadSheet.Tests
{
    public class ExcelExportLayoutTests
    {
        [Fact]
        public void ProcurementHeaders_PreserveHorizontalAndVerticalMergedRanges()
        {
            var document = CreateDocument();
            var sheet = document.AddSheet("Iepirkumi");
            sheet[3, 0].Value = "Reģ. numurs";
            sheet[3, 1].Value = "Pretendenti";
            sheet[4, 1].Value = "Nosaukums";
            sheet[4, 2].Value = "Piedāvājums";
            sheet[4, 3].Value = "Cena";
            sheet.GetCellRange(3, 0, 4, 0).Merge();
            sheet.GetCellRange(3, 1, 3, 3).Merge();
            sheet.GetCellRange(3, 0, 4, 3).SetAlignment(CellHorizontalAlignment.Centred, CellVerticalAlignment.Centred);
            sheet.GetCellRange(3, 0, 4, 3).DrawBorder(Color.Black, BorderLineStyle.Medium);
            sheet[5, 0].Value = "I-001";
            sheet[5, 1].Value = "SIA Ābele";
            sheet[6, 1].Value = "SIA Bērziņš";
            sheet.GetCellRange(5, 0, 6, 0).Merge();

            using var workbook = SaveAndRead(document);
            var result = workbook.GetSheetAt(0);
            Assert.Equal(3, result.NumMergedRegions);
            Assert.Equal(new[] { "A4:A5", "B4:D4", "A6:A7" },
                Enumerable.Range(0, result.NumMergedRegions).Select(i => result.GetMergedRegion(i).FormatAsString()));
            Assert.Equal("Reģ. numurs", Cell(result, 3, 0).StringCellValue);
            Assert.Equal("Pretendenti", Cell(result, 3, 1).StringCellValue);
            Assert.Equal("Cena", Cell(result, 4, 3).StringCellValue);
            Assert.Equal("I-001", Cell(result, 5, 0).StringCellValue);
            Assert.Equal("SIA Bērziņš", Cell(result, 6, 1).StringCellValue);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void CalculateCellHeight_UsesFallbackFontAndPointSize(bool bold, bool italic)
        {
            var document = CreateDocument();
            var sheet = document.AddSheet();
            sheet.Columns[0].Width = 300;
            sheet[0, 0].Value = "Latviešu teksts: āčēģīķļņšūž";
            var font = sheet[0, 0].Formatting.Font;
            font.Name = "Nix Regression Missing Font";
            font.Weight = (ushort)(bold ? Nix.SpreadSheet.Font.BoldWeight : Nix.SpreadSheet.Font.NormalWeight);
            font.Italic = italic;
            font.Size = 200;
            var nativeFont = font.ToNativeFont();
            Assert.Equal("Liberation Sans", nativeFont.Family.Name);
            Assert.Equal(10, nativeFont.Size);
            Assert.Equal(bold, nativeFont.IsBold);
            Assert.Equal(italic, nativeFont.IsItalic);
            ushort smallHeight = sheet[0, 0].CalculateCellHeight();
            font.Size = 400;
            Assert.True(smallHeight > 0);
            Assert.True(sheet[0, 0].CalculateCellHeight() > smallHeight);

            using var workbook = SaveAndRead(document, $"bold-{bold}_italic-{italic}");
            Assert.Equal("Latviešu teksts: āčēģīķļņšūž", Cell(workbook.GetSheetAt(0), 0, 0).StringCellValue);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CalculateCellHeight_WrapsAtColumnWidth(bool wrap)
        {
            var document = CreateDocument();
            var sheet = document.AddSheet();
            sheet[0, 0].Value = "Teksts vairākās rindās ar šauru kolonnas platumu un vairākiem vārdiem";
            sheet[0, 0].Formatting.WrapTextAtRightBorder = wrap;
            sheet.Columns[0].Width = 400;
            ushort wideHeight = sheet[0, 0].CalculateCellHeight();
            sheet.Columns[0].Width = 80;
            Assert.True(sheet[0, 0].CalculateCellHeight() > wideHeight);

            using var workbook = SaveAndRead(document, $"wrap-{wrap}");
            Assert.Equal("Teksts vairākās rindās ar šauru kolonnas platumu un vairākiem vārdiem", Cell(workbook.GetSheetAt(0), 0, 0).StringCellValue);
        }

        [Theory]
        [InlineData("\n", true, false, 204)]
        [InlineData("\r\n", true, false, 204)]
        [InlineData("\n", true, true, 204)]
        [InlineData("\r\n", true, true, 204)]
        [InlineData("\n", false, false, 153)]
        [InlineData("\r\n", false, false, 153)]
        [InlineData("\n", false, true, 153)]
        [InlineData("\r\n", false, true, 153)]
        public void AutoSizeRowHeight_MeasuresExampleTextWithAndWithoutWrapping(string newline, bool wrap, bool merged, int expectedHeight)
        {
            var document = CreateDocument();
            var sheet = document.AddSheet();
            if (merged)
            {
                sheet.Columns[1].Width = 35;
                sheet.Columns[2].Width = 35;
                sheet.GetCellRange("B4:C4").Merge();
            }
            var cell = sheet["B4"];
            cell.Value = "TestXX sdsd sd sd sd sdsd sdsd sdsddd" + newline
                + "sdfsdf sddf gdffdgdfgfdgdfgdfgfdgfdgg dfgdfg gdfgdfgf sdfsdfdsf sdf";
            cell.Formatting.WrapTextAtRightBorder = wrap;

            sheet.AutoSizeRowHeight();

            Assert.Equal((ushort)expectedHeight, sheet[3].Height);
            using var workbook = SaveAndRead(document, $"{(newline == "\n" ? "lf" : "crlf")}_wrap-{wrap}_merged-{merged}");
            var result = workbook.GetSheetAt(0);
            Assert.Equal(expectedHeight * 0.75f, result.GetRow(3).HeightInPoints);
            Assert.Equal(cell.Value, Cell(result, 3, 1).StringCellValue);
            Assert.Equal(wrap, Cell(result, 3, 1).CellStyle.WrapText);
        }

        [Theory]
        [InlineData("\n")]
        [InlineData("\r\n")]
        public void CalculateCellHeight_FitsExplicitLineBreaks(string newline)
        {
            var document = CreateDocument();
            var sheet = document.AddSheet();
            sheet.Columns[0].Width = 300;
            sheet[0, 0].Value = "Rinda";
            ushort singleHeight = sheet[0, 0].CalculateCellHeight();
            sheet[0, 0].Value = "Rinda" + newline + "Rinda" + newline + "Rinda";
            Assert.True(sheet[0, 0].CalculateCellHeight() >= singleHeight * 2);

            using var workbook = SaveAndRead(document, newline == "\n" ? "lf" : "crlf");
            Assert.Equal("Rinda", Cell(workbook.GetSheetAt(0), 0, 0).StringCellValue.Split('\n')[0].TrimEnd('\r'));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void AutoSizeRowHeight_PreservesMinimumForEmptyCells(string? value)
        {
            var document = CreateDocument();
            var sheet = document.AddSheet();
            sheet[0, 0].Value = value;
            sheet.AutoSizeRowHeight();
            Assert.Equal((ushort)17, sheet[0].Height);

            using var workbook = SaveAndRead(document, value == null ? "null" : "empty");
            Assert.Equal(17 * 0.75f, workbook.GetSheetAt(0).GetRow(0).HeightInPoints);
        }

        [Fact]
        public void AutoSizeRowHeight_OverridesExplicitHeightWhenRequested()
        {
            var document = CreateDocument();
            var sheet = document.AddSheet();
            sheet[0, 0].Value = "Rinda";
            sheet[0].Height = 100;
            sheet.AutoSizeRowHeight(true);
            Assert.True(sheet[0].Height > 0);
            Assert.True(sheet[0].Height < 100);

            using var workbook = SaveAndRead(document);
            Assert.Equal(sheet[0].Height!.Value * 0.75f, workbook.GetSheetAt(0).GetRow(0).HeightInPoints);
        }

        [WindowsInstalledFontTheory]
        [InlineData("Arial", 200, 17)]
        [InlineData("Times New Roman", 200, 17)]
        [InlineData("Calibri", 640 , 54)]
        public void AutoSizeRowHeight_UsesExcelLinePitchForInstalledFonts(string name, int size, int linePitch)
        {
            var document = CreateDocument();
            var sheet = document.AddSheet();
            sheet.Columns[0].Width = 300;
            for (int lines = 1; lines <= 4; lines++)
            {
                var cell = sheet[lines - 1, 0];
                cell.Value = string.Join("\n", Enumerable.Repeat("Rinda", lines));
                cell.Formatting.WrapTextAtRightBorder = true;
                var font = cell.Formatting.Font;
                font.Name = name;
                font.Size = (ushort)size;
                Assert.Equal(name, font.ToNativeFont().Family.Name);
            }

            sheet.AutoSizeRowHeight();

            for (int lines = 1; lines <= 4; lines++)
                Assert.Equal((ushort)(lines * linePitch), sheet[lines - 1].Height);

            using var workbook = SaveAndRead(document, $"{name}_{size}");
            for (int lines = 1; lines <= 4; lines++)
                Assert.Equal(lines * linePitch * 0.75f, workbook.GetSheetAt(0).GetRow(lines - 1).HeightInPoints);
        }

        [Theory]
        [InlineData("Nix Regression Missing Font", "Liberation Sans")]
        public void AutoSizeRowHeight_ScalesLinearlyWithSizeAndWeightForFallbackFont(string name, string family)
        {
            AssertLinePitchScalesWithSizeAndWeight(name, family);
        }

        [WindowsInstalledFontTheory]
        [InlineData("Arial")]
        [InlineData("Times New Roman")]
        [InlineData("Calibri")]
        public void AutoSizeRowHeight_ScalesLinearlyWithSizeAndWeightForInstalledFonts(string name)
        {
            AssertLinePitchScalesWithSizeAndWeight(name, name);
        }

        [Theory]
        [InlineData(200, false)]
        [InlineData(200, true)]
        [InlineData(240, false)]
        [InlineData(280, false)]
        [InlineData(280, true)]
        public void AutoSizeRowHeight_FitsMultilineReportAndPreservesExplicitHeight(int fontSize, bool bold)
        {
            var document = CreateDocument();
            var sheet = document.AddSheet("Pārskats");
            sheet.Columns[0].Width = 180;
            sheet[0, 0].Value = "Virsraksts";
            sheet[0].Height = 60;
            sheet[1, 0].Value = "Viena rinda";
            sheet[2, 0].Value = "Pirmā rinda\nOtrā rinda\nTrešā rinda\nCeturtā rinda";
            sheet[2, 0].Formatting.WrapTextAtRightBorder = true;
            sheet[3, 0].Value = "Apvienots virsraksts\nOtrā rinda\nTrešā rinda";
            sheet.GetCellRange(3, 0, 3, 2).Merge();
            sheet[3, 0].Formatting.WrapTextAtRightBorder = true;
            for (int row = 0; row < 4; row++)
            {
                var font = sheet[row, 0].Formatting.Font;
                font.Size = (ushort)fontSize;
                font.Weight = (ushort)(bold ? Nix.SpreadSheet.Font.BoldWeight : Nix.SpreadSheet.Font.NormalWeight);
            }
            sheet.AutoSizeRowHeight();

            ushort linePitch = sheet[1].Height!.Value;
            Assert.Equal((ushort)60, sheet[0].Height);
            Assert.Equal((ushort)(4 * linePitch), sheet[2].Height);
            Assert.Equal((ushort)(3 * linePitch), sheet[3].Height);
            using var workbook = SaveAndRead(document, $"{fontSize}_{(bold ? "bold" : "regular")}");
            var result = workbook.GetSheetAt(0);
            Assert.Equal(45, result.GetRow(0).HeightInPoints);
            Assert.True(result.GetRow(2).HeightInPoints > result.GetRow(1).HeightInPoints);
            Assert.True(result.GetRow(3).HeightInPoints > result.GetRow(1).HeightInPoints);
            Assert.Equal(sheet[2].Height!.Value * 0.75f, result.GetRow(2).HeightInPoints);
            Assert.Equal(sheet[3].Height!.Value * 0.75f, result.GetRow(3).HeightInPoints);
        }

        private static void AssertLinePitchScalesWithSizeAndWeight(string name, string expectedFamily, [CallerMemberName] string testName = "")
        {
            int previousPitch = 0;
            foreach (int size in new[] { 200, 240, 280, 360 })
            {
                ushort regularPitch = 0;
                foreach (bool bold in new[] { false, true })
                {
                    var document = CreateDocument();
                    var sheet = document.AddSheet();
                    sheet.Columns[0].Width = 300;
                    for (int lines = 1; lines <= 4; lines++)
                    {
                        var cell = sheet[lines - 1, 0];
                        cell.Value = string.Join("\n", Enumerable.Repeat("Rinda", lines));
                        cell.Formatting.WrapTextAtRightBorder = true;
                        var font = cell.Formatting.Font;
                        font.Name = name;
                        font.Size = (ushort)size;
                        font.Weight = (ushort)(bold ? Nix.SpreadSheet.Font.BoldWeight : Nix.SpreadSheet.Font.NormalWeight);
                        Assert.Equal(expectedFamily, font.ToNativeFont().Family.Name);
                    }
                    sheet.AutoSizeRowHeight();

                    ushort pitch = sheet[0].Height!.Value;
                    for (int lines = 2; lines <= 4; lines++)
                        Assert.Equal((ushort)(lines * pitch), sheet[lines - 1].Height);

                    using var workbook = SaveAndRead(document, $"{name}_{size}_{(bold ? "bold" : "regular")}", testName);
                    for (int lines = 1; lines <= 4; lines++)
                        Assert.Equal(lines * pitch * 0.75f, workbook.GetSheetAt(0).GetRow(lines - 1).HeightInPoints);

                    if (bold)
                        Assert.True(pitch >= regularPitch);
                    else
                        regularPitch = pitch;
                }

                Assert.True(regularPitch > previousPitch);
                Assert.True(regularPitch >= size / 20.0 * 96 / 72);
                previousPitch = regularPitch;
            }
        }
    }
}
