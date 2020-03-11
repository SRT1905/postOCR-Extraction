namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Used to perform offset search in area, defined by current state of <see cref="TreeNode"/> instance.
    /// </summary>
    public class OffsetNodeProcessor
    {
        private const int SimilaritySearchThreshold = 5;

        private readonly ConfigField configField;
        private readonly LineMapping lineMapping;
        private bool addToParent;
        private TreeNode node;
        private int searchLevel;

        /// <summary>
        /// Initializes a new instance of the <see cref="OffsetNodeProcessor"/> class.
        /// </summary>
        /// <param name="configField">A source description of search field.</param>
        /// <param name="lineMapping">A mapping between document line and paragraphs on it.</param>
        public OffsetNodeProcessor(ConfigField configField, LineMapping lineMapping)
        {
            this.configField = configField;
            this.lineMapping = lineMapping;
        }

        /// <summary>
        /// Perform offset search starting from provided line.
        /// </summary>
        /// <param name="lineNumber">A starting search line.</param>
        /// <param name="node">An instance of <see cref="TreeNode"/> class, in respect to which offset search is made.</param>
        /// <param name="searchLevel">Current search depth level.</param>
        /// <param name="addToParent">Indicates whether found offset node should be added to provided node parent or to provided node itself.</param>
        public void OffsetSearch(int lineNumber, TreeNode node, int searchLevel, bool addToParent = false)
        {
            this.InitializeFields(node, searchLevel, addToParent);
            foreach (var valuePair in this.GetOffsetLines(lineNumber))
            {
                this.AddOffsetNode(
                        this.addToParent
                            ? this.node.Parent
                            : this.node,
                        Convert.ToInt32(valuePair.Key.Split('|')[0]),
                        valuePair.Value,
                        Convert.ToDecimal(valuePair.Key.Split('|')[1]));
            }
        }

        private void InitializeFields(TreeNode node, int searchLevel, bool addToParent)
        {
            this.node = node;
            this.searchLevel = searchLevel;
            this.addToParent = addToParent;
        }

        private Dictionary<string, string> GetOffsetLines(int lineNumber)
        {
            var foundValuesCollection = new Dictionary<string, string>();
            List<int> keys = this.lineMapping.Keys.ToList();
            int lineIndex = keys.IndexOf(lineNumber);
            for (int searchOffset = 1; searchOffset <= SimilaritySearchThreshold; searchOffset++)
            {
                foreach (int offsetIndex in new[] { lineIndex + searchOffset, lineIndex - searchOffset })
                {
                    if (offsetIndex >= 0 && offsetIndex < this.lineMapping.Count)
                    {
                        int line = keys[offsetIndex];
                        var lineChecker = new LineContentChecker(this.lineMapping[line], this.node.Content.UseSoundex);
                        if (lineChecker.CheckLineContents(Utilities.CreateRegexpObject(this.node.Content.TextExpression), this.node.Content.CheckValue))
                        {
                            foundValuesCollection.Add(
                                string.Join("|", line, lineChecker.ParagraphHorizontalLocation),
                                lineChecker.JoinedMatches);
                        }
                    }
                }
            }

            return foundValuesCollection;
        }

        private void AddOffsetNode(TreeNode node, int offsetIndex, string foundValue, decimal position)
        {
            if (node.Content.Lines.Count(item => item == offsetIndex) >= 2)
            {
                return;
            }

            TreeNode childNode = this.InitializeOffsetChildNode(node, offsetIndex, position);
            childNode.Content.FoundValue = foundValue;
            SearchTree.AddSearchValues(this.configField, childNode, this.searchLevel);
        }

        private TreeNode InitializeOffsetChildNode(TreeNode node, int offsetIndex, decimal position)
        {
            var contentBuilder = new TreeNodeContentBuilder(node.Content).AddLine(offsetIndex)
                                                                         .SetHorizontalParagraph(position)
                                                                         .SetNodeLabel(node.Content.NodeLabel)
                                                                         .SetTextExpression(node.Content.TextExpression);

            if (this.addToParent)
            {
                contentBuilder.SetNodeLabel(node.Children[0].Content.NodeLabel)
                              .SetTextExpression(node.Children[0].Content.TextExpression);
            }

            return node.AddChild(contentBuilder.Build());
        }
    }
}
