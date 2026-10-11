using System.Drawing;
using System.Linq;
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
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ToNativeFont_UsesFallbackForUnspecifiedFamily(string? name)
        {
            var font = new Nix.SpreadSheet.Font { Name = name };
            Assert.Equal("Liberation Sans", font.ToNativeFont().Family.Name);
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
        }

        [Fact]
        public void AutoSizeRowHeight_FitsMultilineReportAndPreservesExplicitHeight()
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
            sheet.AutoSizeRowHeight();

            Assert.Equal((ushort)60, sheet[0].Height);
            Assert.True(sheet[2].Height > sheet[1].Height);
            Assert.True(sheet[3].Height > sheet[1].Height);
            using var workbook = SaveAndRead(document);
            var result = workbook.GetSheetAt(0);
            Assert.Equal(45, result.GetRow(0).HeightInPoints);
            Assert.True(result.GetRow(2).HeightInPoints > result.GetRow(1).HeightInPoints);
            Assert.True(result.GetRow(3).HeightInPoints > result.GetRow(1).HeightInPoints);
        }
    }
}
