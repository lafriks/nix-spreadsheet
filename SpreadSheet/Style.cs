using System;

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Describes how cell and its data will be displayed.
	/// </summary>
    public sealed class Style
    {
        #region Default style
        private static Style defStyle = null;

        public static Style Default
        {
            get
            {
                if (defStyle == null)
                {
                    defStyle = new Style();
                }
                return defStyle;
            }
        }

        internal bool Equals(Style comp)
        {
            if (this.wraptext != comp.wraptext)
                return false;
            return true;
        }

        private bool defaultStyle = true;

        public bool IsDefault()
        {
            return this.defaultStyle;
        }
        #endregion

        #region Wrap text
        private bool wraptext = false;

        /// <summary>
        /// Gets or sets if cell value is wrapped.
        /// </summary>
        public bool WrapText
        {
            get
            {
                return this.wraptext;
            }
            set
            {
                this.wraptext = value;
            }
        }
        #endregion
    }
}
