using System;
using System.Runtime.InteropServices;
using SixLabors.Fonts;
using Bitmap = System.Drawing.Bitmap;
using Graphics = System.Drawing.Graphics;
using FontStyle = SixLabors.Fonts.FontStyle;
using SystemFonts = SixLabors.Fonts.SystemFonts;

namespace Nix.SpreadSheet
{
    internal static class TextMeasurement
    {
        private static readonly Lazy<FontFamily> fallbackFont = new Lazy<FontFamily>(LoadFallbackFont);

        internal static float MeasureHeight(string text, Font font, int width)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                return MeasureWindowsHeight(text, font, width);

            FontFamily family;
            if (!SystemFonts.TryGet(font.Name, out family))
                family = fallbackFont.Value;

            FontStyle style = FontStyle.Regular;
            if (font.Italic)
                style |= FontStyle.Italic;
            if (font.Weight >= Font.BoldWeight)
                style |= FontStyle.Bold;

            var options = new TextOptions(family.CreateFont((float)font.Size / 20, style))
            {
                // Row and column dimensions are pixels at Excel's 96 DPI.
                Dpi = 96,
                WrappingLength = width
            };
            return TextMeasurer.MeasureAdvance(text, options).Height;
        }

        private static float MeasureWindowsHeight(string text, Font font, int width)
        {
            using (var nativeFont = font.ToNativeFont())
            using (var bitmap = new Bitmap(1, 1))
            using (var graphics = Graphics.FromImage(bitmap))
                return graphics.MeasureString(text, nativeFont, width).Height;
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
