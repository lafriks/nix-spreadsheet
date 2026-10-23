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

        private static readonly char[] ParagraphSeparators = { '\n', '\r', '\u0085', '\u2028', '\u2029' };

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
            string[] paragraphs = text.Replace("\r\n", "\n").Split(ParagraphSeparators);
            if (width <= 0)
                return paragraphs.Length;

            int lines = 0;
            foreach (string paragraph in paragraphs)
                lines += CountWrappedLines(paragraph, font, shaper, width, breakWords);

            return lines;
        }

        // Greedy wrap over grapheme clusters; trailing spaces hang past the right border like in Excel.
        private static int CountWrappedLines(string paragraph, SKFont font, SKShaper shaper, int width, bool breakWords)
        {
            float MeasureWidth(int from, int to) => shaper.Shape(paragraph.Substring(from, to - from), font).Width;

            int lines = 1;
            int lineStart = 0;
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

                bool breaksAfter = paragraph[index] == '-' || paragraph[index] == '\u200b';
                bool isWordEnd = breaksAfter || end == paragraph.Length || IsWrappingSpace(paragraph, end);

                if ((breakWords || isWordEnd) && MeasureWidth(lineStart, end) > width)
                {
                    if (lastBreak > lineStart)
                    {
                        lineStart = lastBreak;
                        lines++;
                    }

                    if (breakWords && index > lineStart && MeasureWidth(lineStart, end) > width)
                    {
                        lineStart = index;
                        lines++;
                    }
                }

                if (breaksAfter)
                    lastBreak = end;
            }

            return lines;
        }

        private static bool IsWrappingSpace(string text, int index)
        {
            return char.IsWhiteSpace(text, index) && text[index] != '\u00a0' && text[index] != '\u202f';
        }
    }
}
