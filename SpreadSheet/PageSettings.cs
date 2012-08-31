using System;
using System.Collections.Generic;
using System.Text;
using Nix.SpreadSheet.PageSettings;

namespace Nix.SpreadSheet
{
    /// <summary>
    /// Sheet page settings
    /// </summary>
    public class PageSetting
    {
        /// <summary>
        /// Sheet paper size
        /// </summary>
        public PageSize PapeSize { get; set; }

        /// <summary>
        /// Print on page in landscape mode
        /// </summary>
        public bool Landscape { get; set; }

        /// <summary>
        /// Print pages in columns or in rows.
        /// </summary>
        public bool PrintPagesInColumns { get; set; }

        /// <summary>
        /// Print coloured or black and white
        /// </summary>
        public bool PrintColoured { get; set; }

        /// <summary>
        /// Default print quality or draft quality
        /// </summary>
        public bool PrintQualityDefault { get; set; }

        /// <summary>
        /// Print notes
        /// </summary>
        public bool PrintNotes { get; set; }

        public PageSetting()
        {
            PapeSize = PageSize.Default;
            PrintPagesInColumns = true;
            PrintQualityDefault = true;
            PrintColoured = true;
            Landscape = true;
        }
    }
}
