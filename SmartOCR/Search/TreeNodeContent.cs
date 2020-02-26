namespace SmartOCR
{
    using System;
    using System.Collections.Generic;

    public class TreeNodeContent // TODO: add summary.
    {
        public TreeNodeContent()
        {
            this.Lines = new List<long>();
        }

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
            this.RegExPattern = content.RegExPattern;
            this.Status = content.Status;
            this.CheckValue = content.CheckValue;
            this.FoundValue = content.FoundValue;
            this.ValueType = content.ValueType;
            this.SecondSearchParameter = content.SecondSearchParameter;
            this.FirstSearchParameter = content.FirstSearchParameter;
            this.Lines.Clear();
            this.Lines.AddRange(content.Lines);
        }

        public string CheckValue { get; set; }

        public string FoundValue { get; set; }

        public decimal HorizontalParagraph { get; set; }

        public int SecondSearchParameter { get; set; }

        public int FirstSearchParameter { get; set; }

        public List<long> Lines { get; }

        public string Name { get; set; }

        public string NodeLabel { get; set; }

        public string RegExPattern { get; set; }

        public bool Status { get; set; }

        public string ValueType { get; set; }
    }
}