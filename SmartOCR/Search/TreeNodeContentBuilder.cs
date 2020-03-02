namespace SmartOCR
{
    public class TreeNodeContentBuilder : IBuilder<TreeNodeContent> // TODO: add documentation
    {
        private TreeNodeContent nodeContent;

        public TreeNodeContentBuilder()
        {
            this.nodeContent = new TreeNodeContent();
        }

        public TreeNodeContentBuilder(TreeNodeContent content)
        {
            this.nodeContent = new TreeNodeContent(content);
        }

        /// <inheritdoc/>
        public TreeNodeContent Build()
        {
            return this.nodeContent;
        }

        /// <inheritdoc/>
        public void Reset()
        {
            this.nodeContent = new TreeNodeContent();
        }

        public void ResetLines()
        {
            this.nodeContent.Lines.Clear();
        }

        public void SetCheckValue(string value)
        {
            this.nodeContent.CheckValue = value;
        }

        public void SetFirstSearchParameter(int value)
        {
            this.nodeContent.FirstSearchParameter = value;
        }

        public void SetFoundValue(string value)
        {
            this.nodeContent.FoundValue = value;
        }

        public void SetHorizontalParagraph(decimal value)
        {
            this.nodeContent.HorizontalParagraph = value;
        }

        public void SetName(string value)
        {
            this.nodeContent.Name = value;
        }

        public void SetNodeLabel(string value)
        {
            this.nodeContent.NodeLabel = value;
        }

        public void SetRegExPattern(string value)
        {
            this.nodeContent.RegExPattern = value;
        }

        public void SetSecondSearchParameter(int value)
        {
            this.nodeContent.SecondSearchParameter = value;
        }

        public void SetStatus(bool value)
        {
            this.nodeContent.Status = value;
        }

        public void SetValueType(string value)
        {
            this.nodeContent.ValueType = value;
        }

        public bool TryAddLine(int line)
        {
            if (!this.nodeContent.Lines.Contains(line))
            {
                this.nodeContent.Lines.Add(line);
                return true;
            }

            return false;
        }
    }
}
