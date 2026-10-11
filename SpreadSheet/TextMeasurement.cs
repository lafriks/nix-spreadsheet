using System;
using SixLabors.Fonts;
using FontStyle = SixLabors.Fonts.FontStyle;

namespace Nix.SpreadSheet
{
    internal static class TextMeasurement
    {
        private static readonly Lazy<FontFamily> fallbackFont = new Lazy<FontFamily>(LoadFallbackFont);

        internal static float MeasureHeight(string text, Font font, int width)
        {
            var options = new TextOptions(font.ToNativeFont())
            {
                // Row and column dimensions are pixels at Excel's 96 DPI.
                Dpi = 96,
                WrappingLength = width
            };
            return TextMeasurer.MeasureAdvance(text, options).Height;
        }

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
