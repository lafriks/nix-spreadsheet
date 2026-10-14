using System;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Nix.SpreadSheet;
using Nix.SpreadSheet.Provider;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace SpreadSheet.Tests
{
    // Calibri and Times New Roman ship with Windows only; other platforms would silently use the Liberation Sans fallback.
    public sealed class WindowsInstalledFontTheoryAttribute : TheoryAttribute
    {
        public WindowsInstalledFontTheoryAttribute()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                Skip = "Requires fonts installed with Windows.";
        }
    }

    internal static class ExcelExportTestHelper
    {
        // When set, every saved workbook is also written there as <test name>.xls for manual review.
        private const string OutputDirectoryVariable = "SPREADSHEET_TEST_OUTPUT_DIR";

        public static SpreadSheetDocument CreateDocument() => new SpreadSheetDocument { Locale = new RegionInfo("LV") };

        public static HSSFWorkbook SaveAndRead(
            SpreadSheetDocument document,
            string? variant = null,
            [CallerMemberName] string testName = "")
        {
            using var stream = new MemoryStream();
            document.Save(stream, new XlsFileFormatProvider());
            Assert.True(stream.CanRead);
            Assert.True(stream.CanWrite);
            SaveOutput(stream, variant == null ? testName : $"{testName}_{variant}");
            stream.Position = 0;
            return new HSSFWorkbook(stream);
        }

        private static void SaveOutput(MemoryStream stream, string fileName)
        {
            var directory = Environment.GetEnvironmentVariable(OutputDirectoryVariable);
            if (string.IsNullOrWhiteSpace(directory))
                return;
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, fileName + ".xls"), stream.ToArray());
        }

        public static ICell Cell(ISheet sheet, int row, int column)
        {
            var result = sheet.GetRow(row)?.GetCell(column);
            Assert.NotNull(result);
            return result!;
        }
    }
}
