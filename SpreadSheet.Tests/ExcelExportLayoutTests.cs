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

        [WindowsFact]
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
