using System;
using SixLabors.Fonts;

namespace Nix.SpreadSheet
{
    internal static class TextMeasurement
    {
        // Row and column dimensions are pixels at Excel's 96 DPI.
        private const int PixelsPerInch = 96;
        private const int PointsPerInch = 72;

        internal static float MeasureHeight(string text, Font font, int width, bool breakWords = false)
        {
            var nativeFont = font.ToNativeFont();
            var options = new TextOptions(nativeFont)
            {
                Dpi = PixelsPerInch,
                WrappingLength = width,
                WordBreaking = breakWords ? WordBreaking.BreakWord : WordBreaking.Standard
            };

            // MeasureAdvance heights are about 2 px per line shorter than Excel's line pitch, so derive them from the line count.
            var metrics = nativeFont.FontMetrics.HorizontalMetrics;
            float fontSizeInPixels = nativeFont.Size * PixelsPerInch / PointsPerInch;
            float glyphHeight = (float)Math.Round((metrics.Ascender - metrics.Descender + metrics.LineGap) * fontSizeInPixels / nativeFont.FontMetrics.UnitsPerEm);
            return glyphHeight + (TextMeasurer.CountLines(text, options) - 1) * (glyphHeight + ExcelLineSpacing);
        }

        // Extra pixels Excel adds to every line after the first (Arial 10pt: 15 px glyph height, 17 px line pitch; Calibri 11pt: 18 and 20).
        private const int ExcelLineSpacing = 2;
    }
}
