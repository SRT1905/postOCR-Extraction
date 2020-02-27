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
            for (int i = 0; i < this.tables.Count; i++)
            {
                if (this.TryGetDataFromTable(this.tables[i]))
                {
                    return;
                }
            }
        }

        private static bool ExtractValueFromMatch(TreeNode childNode, string itemByExpressionPosition, Regex regexObject)
        {
            int slashCharIndex = childNode.Content.ValueType.IndexOf("/", StringComparison.OrdinalIgnoreCase) + 1;
            string nestedValueType = GetTableValueType(childNode, slashCharIndex);
            TryToValidateMatch(childNode, new MatchProcessor(itemByExpressionPosition, regexObject, nestedValueType));
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
                childNode.Content.FoundValue = matchProcessor.Result;
                childNode.Content.Status = true;
                PropagateStatusInTree(true, childNode);
            }
        }

        private static bool TryToFindMatchInTable(WordTable table, Regex regexObject, string checkValue)
        {
            for (int i = 0; i < table.RowCount; i++)
            {
                for (int j = 0; j < table.ColumnCount; j++)
                {
                    if (table[i, j] != null && ProcessSingleCell(table[i, j], regexObject, checkValue))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ProcessSingleCell(string cellValue, Regex regexObject, string checkValue)
        {
            Match singleMatch = regexObject.Match(cellValue);
            int index = Convert.ToInt32(singleMatch.Groups.Count > 0);
            return new SimilarityDescription(singleMatch.Groups[index].Value, checkValue).CheckStringSimilarity();
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

        private bool TryGetDataFromTable(WordTable table)
        {
            return TryToFindMatchInTable(table, Utilities.CreateRegexpObject(this.tableNode.Content.RegExPattern), this.tableNode.Content.CheckValue)
                ? this.ProcessNodeData(table)
                : false;
        }

        private bool ProcessNodeData(WordTable table)
        {
            TreeNode childNode = this.GetTerminalNode();
            string itemByExpressionPosition = GetCellByNodeContent(table, childNode);
            Regex regexObject = Utilities.CreateRegexpObject(childNode.Content.RegExPattern);
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
