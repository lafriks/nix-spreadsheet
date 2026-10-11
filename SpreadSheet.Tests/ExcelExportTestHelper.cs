using System.Globalization;
using System.IO;
using Nix.SpreadSheet;
using Nix.SpreadSheet.Provider;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace SpreadSheet.Tests
{
    internal static class ExcelExportTestHelper
    {
        public static SpreadSheetDocument CreateDocument() => new SpreadSheetDocument { Locale = new RegionInfo("LV") };

        public static HSSFWorkbook SaveAndRead(SpreadSheetDocument document)
        {
            using var stream = new MemoryStream();
            document.Save(stream, new XlsFileFormatProvider());
            Assert.True(stream.CanRead);
            Assert.True(stream.CanWrite);
            stream.Position = 0;
            return new HSSFWorkbook(stream);
        }

        public static ICell Cell(ISheet sheet, int row, int column)
        {
            var result = sheet.GetRow(row)?.GetCell(column);
            Assert.NotNull(result);
            return result!;
        }
    }
}
