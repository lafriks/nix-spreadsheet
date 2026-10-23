using System;
using System.Globalization;
using SkiaSharp;
using SkiaSharp.HarfBuzz;

namespace Nix.SpreadSheet
{
    internal static class TextMeasurement
    {
        // Row and column dimensions are pixels at Excel's 96 DPI.
        private const int PixelsPerInch = 96;
        private const int PointsPerInch = 72;
        
        // Extra pixels Excel adds to every line after the first (Arial 10pt: 15 px glyph height, 17 px line pitch; Calibri 11pt: 18 and 20).
        private const int ExcelLineSpacing = 2;

        internal static float MeasureHeight(string text, Font font, int width, bool breakWords = false)
        {
            using (var nativeFont = font.ToNativeFont())
            {
                nativeFont.Size *= (float)PixelsPerInch / PointsPerInch;
                nativeFont.LinearMetrics = true;
                nativeFont.Subpixel = true;
                var metrics = nativeFont.Metrics;
                float glyphHeight = (float)Math.Round(metrics.Descent - metrics.Ascent + metrics.Leading);
                using (var shaper = new SKShaper(nativeFont.Typeface))
                    return glyphHeight + (CountLines(text, nativeFont, shaper, width, breakWords) - 1) * (glyphHeight + ExcelLineSpacing);
            }
        }

        private static int CountLines(string text, SKFont font, SKShaper shaper, int width, bool breakWords)
        {
            int lines = 0;
            foreach (string paragraph in text.Replace("\r\n", "\n").Replace('\r', '\n')
                .Replace('\u0085', '\n').Replace('\u2028', '\n').Replace('\u2029', '\n').Split('\n'))
            {
                lines++;
                if (width <= 0)
                    continue;

                int start = 0;
                int lastBreak = 0;
                var elements = StringInfo.GetTextElementEnumerator(paragraph);
                while (elements.MoveNext())
                {
                    int index = elements.ElementIndex;
                    int end = index + elements.GetTextElement().Length;
                    if (IsWrappingSpace(paragraph, index))
                    {
                        lastBreak = end;
                        continue;
                    }

                    bool wordEnd = paragraph[index] == '-' || paragraph[index] == '\u200b';
                    if (!breakWords && !wordEnd && end < paragraph.Length && !IsWrappingSpace(paragraph, end))
                        continue;

                    if (shaper.Shape(paragraph.Substring(start, end - start), font).Width > width)
                    {
                        if (lastBreak > start)
                        {
                            start = lastBreak;
                            lines++;
                        }

                        if (breakWords && index > start && shaper.Shape(paragraph.Substring(start, end - start), font).Width > width)
                        {
                            start = index;
                            lines++;
                        }
                    }

                    if (wordEnd)
                        lastBreak = end;
                }
            }
            return lines;
        }

        private static bool IsWrappingSpace(string text, int index)
        {
            return char.IsWhiteSpace(text, index) && text[index] != '\u00a0' && text[index] != '\u202f';
        }
    }
}
