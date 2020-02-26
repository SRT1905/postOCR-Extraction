namespace SmartOCR
{
    public class TreeNodeContentBuilder : IBuilder<TreeNodeContent>
    {
        #region Fields
        private TreeNodeContent nodeContent;
        #endregion

        #region Constructors
        public TreeNodeContentBuilder()
        {
            nodeContent = new TreeNodeContent();
        }
        public TreeNodeContentBuilder(TreeNodeContent content)
        {
            nodeContent = new TreeNodeContent(content);
        }
        #endregion

        #region Public methods
        public TreeNodeContent Build()
        {
            return nodeContent;
        }
        public void Reset()
        {
            nodeContent = new TreeNodeContent();
        }
        public void ResetLines()
        {
            nodeContent.Lines.Clear();
        }
        public void SetCheckValue(string value)
        {
            nodeContent.CheckValue = value;
        }
        public void SetFirstSearchParameter(int value)
        {
            nodeContent.FirstSearchParameter = value;
        }
        public void SetFoundValue(string value)
        {
            nodeContent.FoundValue = value;
        }
        public void SetHorizontalParagraph(decimal value)
        {
            nodeContent.HorizontalParagraph = value;
        }
        public void SetName(string value)
        {
            nodeContent.Name = value;
        }
        public void SetNodeLabel(string value)
        {
            nodeContent.NodeLabel = value;
        }
        public void SetRegExPattern(string value)
        {
            nodeContent.RegExPattern = value;
        }
        public void SetSecondSearchParameter(int value)
        {
            nodeContent.SecondSearchParameter = value;
        }
        public void SetStatus(bool value)
        {
            nodeContent.Status = value;
        }
        public void SetValueType(string value)
        {
            nodeContent.ValueType = value;
        }
        public bool TryAddLine(long line)
        {
            if (!nodeContent.Lines.Contains(line))
            {
                nodeContent.Lines.Add(line);
                return true;
            }
            return false;
        }
        #endregion
    }
}
