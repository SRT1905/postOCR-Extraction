namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to search values, specified by configuration fields, in Word tables.
    /// </summary>
    public class TableNodeProcessor
    {
        private readonly List<WordTable> tables;
        private readonly TreeNode tableNode;

        /// <summary>
        /// Initializes a new instance of the <see cref="TableNodeProcessor"/> class.
        /// </summary>
        /// <param name="tableCollection">A collection of formatted Word tables.</param>
        /// <param name="fieldNode">A search node, containing information about search field.</param>
        public TableNodeProcessor(List<WordTable> tableCollection, TreeNode fieldNode)
        {
            this.tables = tableCollection;
            this.tableNode = fieldNode;
        }

        /// <summary>
        /// Loops through every Word table searching for specified values.
        /// </summary>
        public void Process()
        {
            foreach (WordTable table in this.tables)
            {
                if (this.TryGetDataFromTable(table))
                {
                    return;
                }
            }

            Utilities.Debug($"No matches were found for table node {this.tableNode.Content.Name}", 3);
        }

        private static bool ExtractValueFromMatch(TreeNode childNode, string itemByExpressionPosition, Regex regexObject)
        {
            TryToValidateMatch(
                childNode,
                new MatchProcessor(
                    itemByExpressionPosition,
                    regexObject,
                    GetTableValueType(
                        childNode,
                        childNode.Content.ValueType.IndexOf("/", StringComparison.OrdinalIgnoreCase) + 1)));
            return true;
        }

        private static string GetTableValueType(TreeNode childNode, int slashCharIndex)
        {
            return slashCharIndex == 0
                ? "String"
                : childNode.Content.ValueType.Substring(slashCharIndex);
        }

        private static void TryToValidateMatch(TreeNode childNode, MatchProcessor matchProcessor)
        {
            if (!string.IsNullOrEmpty(matchProcessor.Result))
            {
                Utilities.Debug($"Match, extracted from table: {matchProcessor.Result}.", 5);
                ValidateMatch(childNode, matchProcessor);
            }
        }

        private static void ValidateMatch(TreeNode childNode, MatchProcessor matchProcessor)
        {
            childNode.Content.FoundValue = matchProcessor.Result;
            childNode.Content.Status = true;
            PropagateStatusInTree(true, childNode);
        }

        private static bool ProcessSingleCell(string cellValue, Regex regexObject, string checkValue)
        {
            Match singleMatch = regexObject.Match(cellValue);
            return new SimilarityDescription(
                singleMatch.Groups[Convert.ToInt32(singleMatch.Groups.Count > 0)].Value,
                checkValue).AreStringsSimilar();
        }

        private static void PropagateStatusInTree(bool status, TreeNode node)
        {
            TreeNode tempNode = node;
            while (tempNode.Parent.Content.Name != "root")
            {
                tempNode = tempNode.Parent;
                tempNode.Content.Status = status;
            }
        }

        private static string GetCellByNodeContent(WordTable table, TreeNode childNode)
        {
            return table[childNode.Content.FirstSearchParameter, childNode.Content.SecondSearchParameter];
        }

        private bool TryToFindMatchInTable(WordTable table, Regex regexObject, string checkValue)
        {
            for (int i = 0; i < table.RowCount; i++)
            {
                for (int j = 0; j < table.ColumnCount; j++)
                {
                    if (table[i, j] != null && ProcessSingleCell(table[i, j], regexObject, checkValue))
                    {
                        Utilities.Debug($"Found match for table node {this.tableNode.Content.Name}. Table anchor: '{table.Anchor}', cell: '{table[i, j]}'.", 3);
                        return true;
                    }
                }
            }

            return false;
        }

        private bool TryGetDataFromTable(WordTable table)
        {
            return this.TryToFindMatchInTable(table, Utilities.CreateRegexpObject(this.tableNode.Content.TextExpression), this.tableNode.Content.CheckValue)
                ? this.ProcessNodeData(table)
                : false;
        }

        private bool ProcessNodeData(WordTable table)
        {
            TreeNode childNode = this.GetTerminalNode();
            string itemByExpressionPosition = GetCellByNodeContent(table, childNode);
            Regex regexObject = Utilities.CreateRegexpObject(childNode.Content.TextExpression);
            return regexObject.IsMatch(itemByExpressionPosition)
                ? ExtractValueFromMatch(childNode, itemByExpressionPosition, regexObject)
                : false;
        }

        private TreeNode GetTerminalNode()
        {
            TreeNode childNode = this.tableNode.Children[0];
            while (childNode.Content.NodeLabel != "Terminal" && childNode.Children.Count != 0)
            {
                childNode = childNode.Children[0];
            }

            return childNode;
        }
    }
}
