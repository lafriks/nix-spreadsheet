using System.Drawing;
using Nix.SpreadSheet;
using NPOI.SS.UserModel;
using Xunit;
using static SpreadSheet.Tests.ExcelExportTestHelper;
using SpreadsheetFont = Nix.SpreadSheet.Font;

namespace SpreadSheet.Tests
{
    public class ExcelExportFormattingTests
    {
        [Fact]
        public void ReportTable_PreservesStylesBordersBackgroundAndColumnWidths()
        {
            var document = CreateDocument();
            var sheet = document.AddSheet("Pārskats pēc jautājumiem");
            var headerStyle = new Style
            {
                WrapText = true,
                WrapTextAtRightBorder = true,
                HorizontalAlignment = CellHorizontalAlignment.Centred,
                VerticalAlignment = CellVerticalAlignment.Centred
            };
            headerStyle.Font.Size = 200;
            headerStyle.Font.Weight = SpreadsheetFont.BoldWeight;
            headerStyle.Font.Italic = true;
            headerStyle.Font.Strikeout = true;
            headerStyle.Font.UnderlineStyle = Nix.SpreadSheet.UnderlineStyle.Double;
            headerStyle.Font.CharSet = CharSet.Baltic;
            sheet[4, 0].Value = "Jautājums";
            sheet[4, 1].Value = 2025;
            sheet[4, 2].Value = 2026;
            sheet.GetCellRange(4, 0, 4, 2).SetFormatting(headerStyle);
            sheet[5, 0].Value = "Iesniegumi";
            sheet[5, 1].Value = 3;
            sheet[5, 2].Value = 0;
            sheet[6, 0].Value = "Kopā";
            sheet[6, 1].Value = 3L;
            sheet[6, 2].Value = 0L;
            sheet.GetCellRange(6, 0, 6, 2).SetBackground(Color.LightGray);
            sheet.GetCellRange(4, 0, 6, 2).DrawTable(Color.Black, BorderLineStyle.Thin, BorderLineStyle.Medium);
            sheet.GetCellRange(4, 0, 4, 2).DrawBorder(Color.Black, BorderLineStyle.Medium);
            sheet.Columns[0].Width = 250;
            sheet.Columns[1].Width = 70;
            sheet.Columns[2].Width = 85;

            for (var row = 4; row <= 6; row++)
                for (var column = 0; column <= 2; column++)
                    sheet[row, column].Formatting.Font.Name = "Times New Roman";

            using var workbook = SaveAndRead(document);
            var result = workbook.GetSheetAt(0);
            var header = Cell(result, 4, 1).CellStyle;
            Assert.True(header.WrapText);
            Assert.Equal(HorizontalAlignment.Center, header.Alignment);
            Assert.Equal(VerticalAlignment.Center, header.VerticalAlignment);
            var font = workbook.GetFontAt(header.FontIndex);
            Assert.Equal("Times New Roman", font.FontName);
            Assert.Equal(10, font.FontHeightInPoints);
            Assert.True(font.IsBold);
            Assert.True(font.IsItalic);
            Assert.True(font.IsStrikeout);
            Assert.Equal(FontUnderlineType.Double, font.Underline);
            Assert.Equal(186, System.Convert.ToInt32(font.Charset));
            Assert.False(workbook.GetFontAt(Cell(result, 5, 0).CellStyle.FontIndex).IsBold);
            Assert.Equal(BorderStyle.Medium, header.BorderTop);
            Assert.Equal(BorderStyle.Medium, header.BorderBottom);
            var body = Cell(result, 5, 1).CellStyle;
            Assert.Equal(BorderStyle.Thin, body.BorderRight);
            Assert.Equal(BorderStyle.Thin, body.BorderBottom);
            Assert.Equal(BorderStyle.Medium, Cell(result, 5, 0).CellStyle.BorderLeft);
            Assert.Equal(BorderStyle.Medium, Cell(result, 5, 2).CellStyle.BorderRight);
            var total = Cell(result, 6, 1).CellStyle;
            Assert.Equal(BorderStyle.Medium, total.BorderBottom);
            Assert.Equal(FillPattern.SolidForeground, total.FillPattern);
            Assert.Equal(new byte[] { 211, 211, 211 },
                workbook.GetCustomPalette().GetColor(total.FillForegroundColor).RGB);
            Assert.Equal(9143, result.GetColumnWidth(0));
            Assert.Equal(2560, result.GetColumnWidth(1));
            Assert.Equal(3109, result.GetColumnWidth(2));
            Assert.Equal(3, Cell(result, 6, 1).NumericCellValue);
            Assert.Equal(0, Cell(result, 5, 2).NumericCellValue);
        }
    }
}
