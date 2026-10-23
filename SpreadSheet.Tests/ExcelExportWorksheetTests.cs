using System.Drawing;
using System.Linq;
using Nix.SpreadSheet;
using Xunit;
using static SpreadSheet.Tests.ExcelExportTestHelper;

namespace SpreadSheet.Tests
{
    public class ExcelExportWorksheetTests
    {
        [Fact]
        public void GroupedExport_PreservesSheetOrderNamesAndIndependentContents()
        {
            var document = CreateDocument();
            var first = document.AddSheet("1 Finanšu nodaļa");
            var second = document.AddSheet("2 Izglītības nodaļa");
            var details = document.AddSheet("Dokumentu saraksts");
            first[0, 0].Value = "Finanšu nodaļa";
            second[0, 0].Value = "Izglītības nodaļa";
            first[4, 0].Value = "Kopīgs virsraksts";
            second[4, 0].Value = "Kopīgs virsraksts";
            first[5, 0].Value = 12;
            second[5, 0].Value = 7;
            details[6, 0].Value = "Saturs";
            details[7, 0].Value = "Tikai detalizētajā sarakstā";
            details.PageSettings.Landscape = false;

            using var workbook = SaveAndRead(document);
            Assert.Equal(3, workbook.NumberOfSheets);
            Assert.Equal(new[] { "1 Finanšu nodaļa", "2 Izglītības nodaļa", "Dokumentu saraksts" },
                Enumerable.Range(0, workbook.NumberOfSheets).Select(workbook.GetSheetName));
            Assert.Equal("Finanšu nodaļa", Cell(workbook.GetSheetAt(0), 0, 0).StringCellValue);
            Assert.Equal("Izglītības nodaļa", Cell(workbook.GetSheetAt(1), 0, 0).StringCellValue);
            Assert.Equal(12, Cell(workbook.GetSheetAt(0), 5, 0).NumericCellValue);
            Assert.Equal(7, Cell(workbook.GetSheetAt(1), 5, 0).NumericCellValue);
            Assert.Equal("Tikai detalizētajā sarakstā", Cell(workbook.GetSheetAt(2), 7, 0).StringCellValue);
            Assert.True(workbook.GetSheetAt(0).PrintSetup.Landscape);
            Assert.False(workbook.GetSheetAt(2).PrintSetup.Landscape);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void EmptyGroupedExport_WithFallbackSheetProducesReadableWorkbook(bool includeHeaders)
        {
            // An export with no groups still needs a worksheet to produce a readable workbook.
            var document = CreateDocument();
            var sheet = document.AddSheet("Sheet1");
            if (includeHeaders)
            {
                sheet[0, 0].Value = "Dokumentu saraksts";
                sheet[3, 0].Value = "Reģ. numurs";
                sheet[3, 1].Value = "Saturs";
                sheet.GetCellRange(3, 0, 3, 1).DrawTable(Color.Black, BorderLineStyle.Thin, BorderLineStyle.Medium);
            }

            using var workbook = SaveAndRead(document, $"includeHeaders-{includeHeaders}");
            Assert.Equal(1, workbook.NumberOfSheets);
            Assert.Equal("Sheet1", workbook.GetSheetName(0));
            if (includeHeaders)
                Assert.Equal("Saturs", Cell(workbook.GetSheetAt(0), 3, 1).StringCellValue);
            else
                Assert.Equal(0, workbook.GetSheetAt(0).PhysicalNumberOfRows);
        }
    }
}
