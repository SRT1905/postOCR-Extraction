namespace SmartOCR.Search.NodeProcessors
{
    using System.Collections.Generic;
    using System.Linq;
    using SmartOCR.Config;
    using SmartOCR.Parse;
    using SmartOCR.Word;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

    /// <summary>
    /// Used to search values, specified by configuration fields, in Word paragraphs, distributed by lines.
    /// </summary>
    public class LineNodeProcessor
    {
        private readonly ConfigField configField;
        private readonly LineMapping lineMapping;

        /// <summary>
        /// Initializes a new instance of the <see cref="LineNodeProcessor"/> class.
        /// </summary>
        /// <param name="configField">A source description of search field.</param>
        /// <param name="lineMapping">A mapping between document line and paragraphs on it.</param>
        public LineNodeProcessor(ConfigField configField, LineMapping lineMapping)
        {
            this.configField = configField;
            this.lineMapping = lineMapping;
        }

        /// <summary>
        /// Loops through every Word paragraph searching for specified values.
        /// </summary>
        /// <param name="lineNode">An instance of <see cref="TreeNode"/> class, in respect to which search is made.</param>
        /// <param name="searchLevel">Level of search, equals 0 by default.</param>
        public void ProcessLineNode(TreeNode lineNode, int searchLevel = 0)
        {
            if (lineNode.Content.NodeLabel != "Terminal")
            {
                this.ProcessNonTerminalNode(lineNode, searchLevel);
            }
            else
            {
                new TerminalNodeProcessor(this.configField, this.lineMapping).Process(lineNode, searchLevel);
            }
        }

        private static bool CheckLineContents(TreeNodeContent lineNodeContent, LineContentChecker lineChecker)
        {
            return lineChecker.CheckLineContents(
                Utilities.CreateRegexpObject(lineNodeContent.TextExpression),
                lineNodeContent.CheckValue);
        }

        private static bool GetLineCheckStatus(TreeNodeContent lineNodeContent, int lineNumber, LineContentChecker lineChecker)
        {
            bool checkStatus = CheckLineContents(lineNodeContent, lineChecker);

            if (checkStatus)
            {
                lineNodeContent.HorizontalParagraph = lineChecker.ParagraphHorizontalLocation;
                Utilities.Debug($"Found match in line {lineNumber}: {lineChecker.JoinedMatches}", 4);
            }

            lineNodeContent.FoundValue = lineChecker.JoinedMatches;
            return checkStatus;
        }

        private void ProcessNonTerminalNode(TreeNode lineNode, int searchLevel)
        {
            int lineIndex = 0;
            while (lineIndex < lineNode.Content.Lines.Count)
            {
                this.ProcessSingleLineNode(lineNode, lineNode.Content.Lines[lineIndex], searchLevel);
                lineIndex++;
            }
        }

        private void ProcessSingleLineNode(TreeNode lineNode, int lineNumber, int searchLevel)
        {
            Utilities.Debug($"Processing node '{lineNode.Content.Name}' labeled as '{lineNode.Content.NodeLabel}'.", 3);
            if (this.GetLineStatus(lineNode.Content, lineNumber))
            {
                this.ProcessNextLevelNodes(lineNode, lineNumber, searchLevel);
            }
            else
            {
                this.DoOffsetSearch(lineNode, lineNumber, searchLevel);
            }
        }

        private void DoOffsetSearch(TreeNode lineNode, int lineNumber, int searchLevel)
        {
            new OffsetNodeProcessor(this.configField, this.lineMapping).OffsetSearch(lineNode, lineNumber, searchLevel, addToParentStatus: true);
        }

        private void ProcessNextLevelNodes(TreeNode lineNode, int lineNumber, int searchLevel)
        {
            this.SetOffsetChildrenLines(lineNode, lineNumber);
            this.ProcessLineNodeChildren(lineNode, searchLevel);
        }

        private bool GetLineStatus(TreeNodeContent lineNodeContent, int lineNumber)
        {
            return this.lineMapping.ContainsKey(lineNumber) && this.TryMatchLineData(lineNodeContent, lineNumber);
        }

        private void ProcessLineNodeChildren(TreeNode lineNode, int searchLevel)
        {
            int childIndex = 0;
            while (childIndex < lineNode.Children.Count)
            {
                this.ProcessLineNode(lineNode.Children[childIndex], searchLevel + 1);
                childIndex++;
            }
        }

        private void SetOffsetChildrenLines(TreeNode node, int line)
        {
            for (int i = 0; i < node.Children.Count; i++)
            {
                this.SetOffsetLineForSingleChild(node, line, i);
            }
        }

        private void SetOffsetLineForSingleChild(TreeNode node, int line, int childIndex)
        {
            TreeNodeContent childContent = node.Children[childIndex].Content;
            childContent.HorizontalParagraph = node.Content.HorizontalParagraph;
            List<int> keys = this.lineMapping.Keys.ToList();
            this.TryAddLineToChildContent(
                childContent,
                keys,
                keys.IndexOf(line) + childContent.FirstSearchParameter);
        }

        private void TryAddLineToChildContent(TreeNodeContent childContent, List<int> keys, int lineIndex)
        {
            if (lineIndex < 0 || lineIndex >= this.lineMapping.Count)
            {
                return;
            }

            childContent.Lines.Clear();
            childContent.Lines.Add(keys[lineIndex]);
        }

        private bool TryMatchLineData(TreeNodeContent lineNodeContent, int lineNumber)
        {
            return GetLineCheckStatus(lineNodeContent, lineNumber, new LineContentChecker(this.lineMapping[lineNumber], lineNodeContent));
        }
    }
}
