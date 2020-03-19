namespace SmartOCR
{
    /// <summary>
    /// Used to search values, specified by configuration field.
    /// </summary>
    public class FieldNodeProcessor
    {
        private readonly ConfigField configField;
        private readonly LineMapping lineMapping;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldNodeProcessor"/> class.
        /// </summary>
        /// <param name="configField">A source description of search field.</param>
        /// <param name="lineMapping">A mapping between document line and paragraphs on it.</param>
        public FieldNodeProcessor(ConfigField configField, LineMapping lineMapping)
        {
            this.configField = configField;
            this.lineMapping = lineMapping;
        }

        /// <summary>
        /// Performs field node processing.
        /// </summary>
        /// <param name="fieldNode">A field node instance.</param>
        public void ProcessFieldNode(TreeNode fieldNode)
        {
            fieldNode = this.InitializeFieldNode(fieldNode);

            if (fieldNode.Content.Lines[0] == 0)
            {
                return;
            }

            Utilities.Debug($"Performing search for field node '{fieldNode.Content.Name}' data.", 2);
            this.ProcessFieldNodeChildren(fieldNode);
        }

        private TreeNode InitializeFieldNode(TreeNode fieldNode)
        {
            if (fieldNode.Content.Lines[0] != 0)
            {
                return fieldNode;
            }

            Utilities.Debug($"Initializing field node '{fieldNode.Content.Name}' data.", 2);
            fieldNode = new UndefinedNodeProcessor(
                fieldNode,
                this.lineMapping,
                this.configField).GetProcessedNode();

            return fieldNode;
        }

        private void ProcessFieldNodeChildren(TreeNode fieldNode)
        {
            foreach (var lineNode in fieldNode.Children)
            {
                this.ProcessSingleLineNode(lineNode);
            }
        }

        private void ProcessSingleLineNode(TreeNode lineNode)
        {
            new LineNodeProcessor(this.configField, this.lineMapping).ProcessLineNode(lineNode);
        }
    }
}
