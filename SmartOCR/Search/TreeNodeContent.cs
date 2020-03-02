namespace SmartOCR
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Class represents description of TreeNode search parameters.
    /// </summary>
    public class TreeNodeContent
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TreeNodeContent"/> class.
        /// Properties are initialized with default values.
        /// </summary>
        public TreeNodeContent()
        {
            this.Lines = new List<int>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TreeNodeContent"/> class.
        /// Properties are initialized with values from other <see cref="TreeNodeContent"/> instance properties.
        /// </summary>
        /// <param name="content">A source instance of the <see cref="TreeNodeContent"/> class.</param>
        public TreeNodeContent(TreeNodeContent content)
            : this()
        {
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            this.HorizontalParagraph = content.HorizontalParagraph;
            this.Name = content.Name;
            this.NodeLabel = content.NodeLabel;
            this.TextExpression = content.TextExpression;
            this.Status = content.Status;
            this.CheckValue = content.CheckValue;
            this.FoundValue = content.FoundValue;
            this.ValueType = content.ValueType;
            this.SecondSearchParameter = content.SecondSearchParameter;
            this.FirstSearchParameter = content.FirstSearchParameter;
            this.Lines.Clear();
            this.Lines.AddRange(content.Lines);
        }

        /// <summary>
        /// Gets or sets value to check with regular expression.
        /// </summary>
        public string CheckValue { get; set; }

        /// <summary>
        /// Gets or sets value that is matched by regular expression.
        /// </summary>
        public string FoundValue { get; set; }

        /// <summary>
        /// Gets or sets horizontal location of paragraph where text match was found.
        /// </summary>
        public decimal HorizontalParagraph { get; set; }

        /// <summary>
        /// Gets or sets primary numeric parameter of search expression.
        /// </summary>
        public int FirstSearchParameter { get; set; }

        /// <summary>
        /// Gets or sets secondary numeric parameter of search expression.
        /// </summary>
        public int SecondSearchParameter { get; set; }

        /// <summary>
        /// Gets collection of document lines where match was found.
        /// </summary>
        public List<int> Lines { get; }

        /// <summary>
        /// Gets or sets field name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets TreeNode type.
        /// </summary>
        public string NodeLabel { get; set; }

        /// <summary>
        /// Gets or sets regular expression pattern that is used to match text.
        /// </summary>
        public string TextExpression { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether match has been found.
        /// </summary>
        public bool Status { get; set; }

        /// <summary>
        /// Gets or sets value type of searched value.
        /// </summary>
        public string ValueType { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether <see cref="TextExpression"/> is referred to as Regex pattern or Soundex expression.
        /// </summary>
        public bool UseSoundex { get; set; }
    }
}