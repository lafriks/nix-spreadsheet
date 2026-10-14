using System;
using SixLabors.Fonts;
using FontStyle = SixLabors.Fonts.FontStyle;

namespace Nix.SpreadSheet
{
    internal static class TextMeasurement
    {
        private static readonly Lazy<FontFamily> fallbackFont = new Lazy<FontFamily>(LoadFallbackFont);

        // Row and column dimensions are pixels at Excel's 96 DPI.
        private const int PixelsPerInch = 96;
        private const int PointsPerInch = 72;

        internal static float MeasureHeight(string text, Font font, int width)
        {
            var nativeFont = font.ToNativeFont();
            var options = new TextOptions(nativeFont)
            {
                Dpi = PixelsPerInch,
                WrappingLength = width
            };

            // MeasureAdvance heights are about 2 px per line shorter than Excel's line pitch, so derive them from the line count.
            var metrics = nativeFont.FontMetrics.HorizontalMetrics;
            float fontSizeInPixels = nativeFont.Size * PixelsPerInch / PointsPerInch;
            float glyphHeight = (float)Math.Round((metrics.Ascender - metrics.Descender + metrics.LineGap) * fontSizeInPixels / nativeFont.FontMetrics.UnitsPerEm);
            return glyphHeight + (TextMeasurer.CountLines(text, options) - 1) * (glyphHeight + ExcelLineSpacing);
        }

        // Extra pixels Excel adds to every line after the first (Arial 10pt: 15 px glyph height, 17 px line pitch; Calibri 11pt: 18 and 20).
        private const int ExcelLineSpacing = 2;

        internal static SixLabors.Fonts.Font CreateFont(Font font)
        {
            FontFamily family;
            if (string.IsNullOrEmpty(font.Name) || !SystemFonts.TryGet(font.Name, out family))
                family = fallbackFont.Value;

            FontStyle style = FontStyle.Regular;
            if (font.Italic)
                style |= FontStyle.Italic;
            if (font.Weight >= Font.BoldWeight)
                style |= FontStyle.Bold;

            return family.CreateFont((float)font.Size / 20, style);
        }

        private static FontFamily LoadFallbackFont()
        {
            var collection = new FontCollection();
            var assembly = typeof(TextMeasurement).Assembly;
            foreach (string style in new[] { "Regular", "Bold", "Italic", "BoldItalic" })
            {
                using (var stream = assembly.GetManifestResourceStream("Nix.SpreadSheet.Fonts.LiberationSans-" + style + ".ttf"))
                    collection.Add(stream);
            }
            return collection.Get("Liberation Sans");
        }
    }
}
