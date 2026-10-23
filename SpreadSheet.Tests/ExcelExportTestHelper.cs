using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using Nix.SpreadSheet;
using Nix.SpreadSheet.Provider;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SkiaSharp;
using Xunit;
using Xunit.Sdk;

namespace SpreadSheet.Tests
{
    public sealed class InstalledFontDataAttribute : DataAttribute
    {
        private readonly object[] data;

        public InstalledFontDataAttribute(string name, params object[] values)
        {
            using var style = new SKFontStyle(SKFontStyleWeight.Normal, SKFontStyleWidth.Normal, SKFontStyleSlant.Upright);
            using var typeface = SKFontManager.Default.MatchFamily(name, style);
            if (typeface == null)
                Skip = $"SkiaSharp could not find '{name}' or a system substitute.";

            data = new object[values.Length + 2];
            data[0] = name;
            data[1] = typeface?.FamilyName!;
            Array.Copy(values, 0, data, 2, values.Length);
        }

        public override IEnumerable<object[]> GetData(MethodInfo testMethod)
        {
            yield return data;
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
