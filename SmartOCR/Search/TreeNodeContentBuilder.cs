namespace SmartOCR
{
    /// <summary>
    /// Used to build an instance of <see cref="TreeNodeContent"/> class in stepwise fashion.
    /// </summary>
    public class TreeNodeContentBuilder : IBuilder<TreeNodeContent>
    {
        private TreeNodeContent nodeContent;

        /// <summary>
        /// Initializes a new instance of the <see cref="TreeNodeContentBuilder"/> class.
        /// </summary>
        public TreeNodeContentBuilder()
        {
            this.Reset();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TreeNodeContentBuilder"/> class.
        /// Initial builder state is defined by argument of <see cref="TreeNodeContent"/> type.
        /// </summary>
        /// <param name="content">A base for building a <see cref="TreeNodeContent"/> instance.</param>
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

        /// <summary>
        /// Removes all stored lines, where supposed match is made.
        /// </summary>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder ResetLines()
        {
            this.nodeContent.Lines.Clear();
            return this;
        }

        /// <summary>
        /// Sets value that should be compared with supposed match.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetCheckValue(string value)
        {
            this.nodeContent.CheckValue = value;
            return this;
        }

        /// <summary>
        /// Sets first numeric search parameters.
        /// Could be line offset or table row index, depending on value type.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetFirstSearchParameter(int value)
        {
            this.nodeContent.FirstSearchParameter = value;
            return this;
        }

        /// <summary>
        /// Sets value that is extracted from supposed match.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetFoundValue(string value)
        {
            this.nodeContent.FoundValue = value;
            return this;
        }

        /// <summary>
        /// Sets distance between left margin of document and paragraph, where match was found.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetHorizontalParagraph(decimal value)
        {
            this.nodeContent.HorizontalParagraph = value;
            return this;
        }

        /// <summary>
        /// Sets name of node.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetName(string value)
        {
            this.nodeContent.Name = value;
            return this;
        }

        /// <summary>
        /// Sets search level of node.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetNodeLabel(string value)
        {
            this.nodeContent.NodeLabel = value;
            return this;
        }

        /// <summary>
        /// Sets expression (RegEx or Soundex) that is used to find matches.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetTextExpression(string value)
        {
            this.nodeContent.TextExpression = value;
            return this;
        }

        /// <summary>
        /// Sets second numeric search parameters.
        /// Could be paragraph search status or table column index, depending on value type.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetSecondSearchParameter(int value)
        {
            this.nodeContent.SecondSearchParameter = value;
            return this;
        }

        /// <summary>
        /// Sets indicator whether match was found.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetStatus(bool value)
        {
            this.nodeContent.Status = value;
            return this;
        }

        /// <summary>
        /// Sets node value type.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetValueType(string value)
        {
            this.nodeContent.ValueType = value;
            return this;
        }

        /// <summary>
        /// Tries to add new line to line collection.
        /// </summary>
        /// <param name="line">Line to add.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder AddLine(int line)
        {
            if (!this.nodeContent.Lines.Contains(line))
            {
                this.nodeContent.Lines.Add(line);
            }

            return this;
        }

        /// <summary>
        /// Sets indicator whether Soundex is used to find matches.
        /// </summary>
        /// <param name="value">Value to set.</param>
        /// <returns>An updated instance of <see cref="TreeNodeContentBuilder"/>.</returns>
        public TreeNodeContentBuilder SetSoundexUsageStatus(bool value)
        {
            this.nodeContent.UseSoundex = value;
            return this;
        }
    }
}
