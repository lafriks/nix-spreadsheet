using System.Drawing;
using Nix.SpreadSheet;
using NPOI.SS.UserModel;
using Xunit;
using static SpreadSheet.Tests.ExcelExportTestHelper;
using SpreadsheetFont = Nix.SpreadSheet.Font;

namespace SpreadSheet.Tests
{
    public class ExcelExportCellValueTests
    {
        [Fact]
        public void DocumentList_PreservesLatvianTextAmountsAndOptionalValues()
        {
            var document = CreateDocument();
            var sheet = document.AddSheet("Saņemtie dokumenti");
            sheet[0, 0].Value = "Pašvaldības dokumentu saraksts";
            string[] headers = { "Reģ. numurs", "Saturs", "Summa", "Piezīmes", "Parakstīts" };
            for (var column = 0; column < headers.Length; column++)
            {
                sheet[3, column].Value = headers[column];
                sheet[3, column].Formatting.Font.Weight = SpreadsheetFont.BoldWeight;
                sheet.Columns[column].Width = 150;
            }

            sheet[4, 0].Value = "001/Ā-2026";
            sheet[4, 1].Value = "Jānis Bērziņš\r\nĶekava, Šķūņu iela — ☑";
            sheet[4, 2].Value = 1234.56m;
            sheet[5, 2].Value = -12.34m;
            sheet[6, 2].Value = 0m;
            sheet[4, 3].Value = null;
            sheet[5, 3].Value = "";
            sheet[4, 4].Value = "Jā";
            sheet[5, 4].Value = "Nē";
            sheet.GetCellRange(3, 0, (int)sheet.LastRow, (int)sheet.LastColumn)
                .DrawTable(Color.Black, BorderLineStyle.Thin, BorderLineStyle.Medium);

            using var workbook = SaveAndRead(document);
            var result = workbook.GetSheetAt(0);
            Assert.Equal("Saņemtie dokumenti", result.SheetName);
            Assert.Equal("Pašvaldības dokumentu saraksts", Cell(result, 0, 0).StringCellValue);
            for (var column = 0; column < headers.Length; column++)
                Assert.Equal(headers[column], Cell(result, 3, column).StringCellValue);
            Assert.Equal(CellType.String, Cell(result, 4, 0).CellType);
            Assert.Equal("001/Ā-2026", Cell(result, 4, 0).StringCellValue);
            Assert.Equal("Jānis Bērziņš\r\nĶekava, Šķūņu iela — ☑", Cell(result, 4, 1).StringCellValue);
            Assert.Equal(CellType.Numeric, Cell(result, 4, 2).CellType);
            Assert.Equal(1234.56, Cell(result, 4, 2).NumericCellValue, 8);
            Assert.Equal(-12.34, Cell(result, 5, 2).NumericCellValue, 8);
            Assert.Equal(0, Cell(result, 6, 2).NumericCellValue);
            Assert.Equal(CellType.Blank, Cell(result, 4, 3).CellType);
            Assert.Equal(CellType.Blank, Cell(result, 5, 3).CellType);
            Assert.Equal("Jā", Cell(result, 4, 4).StringCellValue);
            Assert.Equal("Nē", Cell(result, 5, 4).StringCellValue);
        }

        [Fact]
        public void RegistrationDates_RemainTextWithTheRequestedFormat()
        {
            // Formatted date strings must not become numeric Excel serial dates.
            var document = CreateDocument();
            var sheet = document.AddSheet("Reģistrācijas datumi");
            sheet[4, 0].Value = "09.10.2026";
            sheet[4, 0].Formatting.Format = "dd.MM.yyyy";
            sheet[4, 1].Value = "09.10.2026 14:35";
            sheet[4, 1].Formatting.Format = "dd.MM.yyyy HH:mm";
            sheet[5, 0].Value = "";
            sheet[5, 0].Formatting.Format = "dd.MM.yyyy";

            using var workbook = SaveAndRead(document);
            var result = workbook.GetSheetAt(0);
            Assert.Equal(CellType.String, Cell(result, 4, 0).CellType);
            Assert.Equal("09.10.2026", Cell(result, 4, 0).StringCellValue);
            Assert.Equal("dd.MM.yyyy", Cell(result, 4, 0).CellStyle.GetDataFormatString());
            Assert.Equal(CellType.String, Cell(result, 4, 1).CellType);
            Assert.Equal("09.10.2026 14:35", Cell(result, 4, 1).StringCellValue);
            Assert.Equal("dd.MM.yyyy HH:mm", Cell(result, 4, 1).CellStyle.GetDataFormatString());
            Assert.Equal(CellType.Blank, Cell(result, 5, 0).CellType);
            Assert.Equal("dd.MM.yyyy", Cell(result, 5, 0).CellStyle.GetDataFormatString());
        }
    }
}
