using System.Linq;
using Xunit;
using static SpreadSheet.Tests.ExcelExportTestHelper;

namespace SpreadSheet.Tests
{
    public class ExcelExportLargeDataTests
    {
        [Fact]
        public void LargeDocumentList_PreservesEveryRowAndSharedStringAcrossRecordBoundaries()
        {
            var document = CreateDocument();
            var sheet = document.AddSheet("Dokumentu saraksts");
            const int count = 2048;
            sheet[3, 0].Value = "Nr.";
            sheet[3, 1].Value = "Saturs";
            sheet[3, 2].Value = "Struktūrvienība";
            for (var i = 0; i < count; i++)
            {
                sheet[i + 4, 0].Value = i + 1;
                sheet[i + 4, 1].Value = LongDescription(i);
                sheet[i + 4, 2].Value = "Finanšu nodaļa";
            }
            // Footer after a gap also exercises nonconsecutive row blocks.
            sheet[count + 7, 0].Value = "Sagatavoja Jānis Bērziņš";

            using var workbook = SaveAndRead(document);
            var result = workbook.GetSheetAt(0);
            Assert.Equal(count + 7, result.LastRowNum);
            for (var i = 0; i < count; i++)
            {
                Assert.Equal(i + 1, Cell(result, i + 4, 0).NumericCellValue);
                Assert.Equal(LongDescription(i), Cell(result, i + 4, 1).StringCellValue);
                Assert.Equal("Finanšu nodaļa", Cell(result, i + 4, 2).StringCellValue);
            }
            Assert.Equal("Sagatavoja Jānis Bērziņš", Cell(result, count + 7, 0).StringCellValue);
        }

        [Fact]
        public void LongDocumentContent_PreservesOneUnicodeStringAcrossContinuationRecords()
        {
            var document = CreateDocument();
            var sheet = document.AddSheet("Dokumenta saturs");
            var content = string.Concat(Enumerable.Repeat("Ābele, ķirši, Šķūņu iela.\r\n", 600));
            sheet[4, 0].Value = content;
            sheet[4, 0].Formatting.WrapTextAtRightBorder = true;
            sheet[5, 0].Value = "Nākamais dokuments";

            using var workbook = SaveAndRead(document);
            var result = workbook.GetSheetAt(0);
            Assert.Equal(content, Cell(result, 4, 0).StringCellValue);
            Assert.True(Cell(result, 4, 0).CellStyle.WrapText);
            Assert.Equal("Nākamais dokuments", Cell(result, 5, 0).StringCellValue);
        }

        private static string LongDescription(int index) =>
            $"Dokuments {index:D5}: " + string.Concat(Enumerable.Repeat("Pašvaldības lēmums par izglītību un īpašumu. ", 12));
    }
}
