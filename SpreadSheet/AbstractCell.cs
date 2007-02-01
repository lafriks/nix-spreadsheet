using System;

namespace Nix.SpreadSheet
{
	/// <summary>
	/// Base class for one or more cells.
	/// </summary>
	public abstract class AbstractCell
	{
        #region Comment
        private Comment comment = null;

        /// <summary>
        /// Gets or sets cell comment.
        /// </summary>
        public Comment Comment
        {
            get
            {
                if (this.comment == null)
                    this.comment = new Comment();
                return this.comment;
            }
            set
            {
                this.comment = value;
            }
        }
        #endregion

        #region Style
        private Style style = null;

        /// <summary>
        /// Gets or sets cell style.
        /// </summary>
        public Style Style
        {
            get
            {
                if (this.style == null)
                    this.style = new Style();
                return this.style;
            }
            set
            {
                this.style = value;
            }
        }

        /// <summary>
        /// Chack if cell style is equal to default style.
        /// </summary>
        /// <returns>True if cell style is same as default.</returns>
        public bool IsDefaultStyle()
        {
            return this.style.IsDefault();
        }
        #endregion

        #region Formula
        private string formula = string.Empty;

        public string Formula
        {
            get
            {
                return this.formula;
            }
            set
            {
                this.formula = value;
            }
        }
        #endregion

        #region Value
        private object val;

        /// <summary>
        /// Gets or sets value on one or more cells. When setting value formula is removed.
        /// </summary>
        public object Value
        {
            get
            {
                // TODO: Evalute formula to get real value
                if (this.formula != string.Empty)
                    return this.formula;
                else
                    return this.val;
            }
            set
            {
                this.formula = string.Empty;
                this.val = value;
            }
        }
        #endregion
	}
}
