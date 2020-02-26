namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;

    public class TableNodeProcessor
    {
        private readonly List<WordTable> tables;
        private readonly TreeNode tableNode;

        public TableNodeProcessor(List<WordTable> tableCollection, TreeNode fieldNode)
        {
            this.tables = tableCollection;
            this.tableNode = fieldNode;
        }

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
            string nestedValueType = slashCharIndex == 0
                ? "String"
                : childNode.Content.ValueType.Substring(slashCharIndex);
            var matchProcessor = new MatchProcessor(itemByExpressionPosition, regexObject, nestedValueType);
            if (!string.IsNullOrEmpty(matchProcessor.Result))
            {
                childNode.Content.FoundValue = matchProcessor.Result;
                childNode.Content.Status = true;
                PropagateStatusInTree(true, childNode);
            }

            return true;
        }

        private static bool TryToFindMatchInTable(WordTable table, Regex regexObject, string checkValue)
        {
            for (int i = 0; i < table.RowCount; i++)
            {
                for (int j = 0; j < table.ColumnCount; j++)
                {
                    if (table[i, j] != null && regexObject.IsMatch(table[i, j]))
                    {
                        Match singleMatch = regexObject.Match(table[i, j]);
                        int index = Convert.ToInt32(singleMatch.Groups.Count > 0);
                        var similarity = new SimilarityDescription(singleMatch.Groups[index].Value, checkValue);

                        if (similarity.CheckStringSimilarity())
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
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

        private bool TryGetDataFromTable(WordTable table)
        {
            if (!TryToFindMatchInTable(
                table,
                Utilities.CreateRegexpObject(this.tableNode.Content.RegExPattern),
                this.tableNode.Content.CheckValue))
            {
                return false;
            }

            return this.ProcessNodeData(table);
        }

        private bool ProcessNodeData(WordTable table)
        {
            TreeNode childNode = this.GetTerminalNode();

            string itemByExpressionPosition = table[childNode.Content.FirstSearchParameter, childNode.Content.SecondSearchParameter];
            Regex regexObject = Utilities.CreateRegexpObject(childNode.Content.RegExPattern);
            if (regexObject.IsMatch(itemByExpressionPosition))
            {
                return ExtractValueFromMatch(childNode, itemByExpressionPosition, regexObject);
            }

            return false;
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
