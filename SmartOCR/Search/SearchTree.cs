using System;
using System.Collections.Generic;

namespace SmartOCR
{
    public class SearchTree // TODO: add summary.
    {
        #region Fields
        private readonly ConfigData ConfigData;
        private TreeNode treeStructure;
        #endregion

        #region Properties
        public List<TreeNode> Children
        {
            get
            {
                return treeStructure.Children;
            }
        }
        #endregion

        #region Constructors
        public SearchTree(ConfigData configData)
        {
            this.ConfigData = configData;
        }
        #endregion

        #region Public methods
        public Dictionary<string, string> GetValuesFromTree()
        {
            var finalValues = new Dictionary<string, string>();
            foreach (ConfigField field in ConfigData.Fields)
            {
                List<string> childrenCollection = GetChildrenByFieldName(field.Name);
                var result = new HashSet<string>();
                if (childrenCollection.Count != 0)
                {
                    result.UnionWith(childrenCollection);
                }
                else
                {
                    var preTerminalCollection = GetDataFromPreTerminalNodes(field.Name);
                    result.UnionWith(preTerminalCollection);
                }
                finalValues.Add(field.Name, string.Join("|", result));
            }
            return finalValues;
        }
        public void PopulateTree()
        {
            TreeNode root = TreeNode.CreateRoot();
            foreach (ConfigField field in ConfigData.Fields)
            {
                TreeNode fieldNode = AddFieldNode(root, field);
                AddSearchValues(field, fieldNode);
            }
            treeStructure = root;
        }
        #endregion

        #region Private methods

        private List<string> GetChildrenByFieldName(string fieldName)
        {
            var childrenCollection = new List<string>();
            foreach (TreeNode fieldNode in treeStructure.Children)
            {
                if (fieldNode.Content.Name == fieldName)
                {
                    GetNodeChildren(fieldNode, childrenCollection);
                    break;
                }
            }

            return childrenCollection;
        }
        private void GetDataFromNode(TreeNode node, Dictionary<bool, HashSet<string>> foundData)
        {
            if (node.Children.Count == 0)
            {
                return;
            }

            foreach (TreeNode child in node.Children)
            {
                GetDataFromNode(child, foundData);
            }
            if (!string.IsNullOrEmpty(node.Content.FoundValue))
            {
                if (node.Content.Status)
                {
                    foundData[true].Add(node.Content.FoundValue);
                }
                else
                {
                    if (node.Parent.Content.NodeLabel != "Field")
                    {
                        foundData[false].Add(node.Content.FoundValue);
                    }
                }

            }
        }
        private HashSet<string> GetDataFromPreTerminalNodes(string fieldName)
        {
            var foundData = new Dictionary<bool, HashSet<string>>()
            {
                { true, new HashSet<string>() },
                { false, new HashSet<string>() }
            };
            foreach (TreeNode node in treeStructure.Children)
            {
                if (node.Content.Name == fieldName)
                {
                    GetDataFromNode(node, foundData);
                    if (foundData[true].Count != 0)
                    {
                        return foundData[true];
                    }
                    return foundData[false];

                }
            }
            return new HashSet<string>();
        }
        private void GetNodeChildren(TreeNode node, List<string> childrenCollection)
        {
            if (node.Children.Count == 0)
            {
                if (node.Content.Status)
                {
                    childrenCollection.Add(node.Content.FoundValue);
                }
                return;
            }

            foreach (TreeNode child in node.Children)
            {
                GetNodeChildren(child, childrenCollection);
            }
        }
        #endregion

        #region Public static methods
        public static void AddSearchValues(ConfigField fieldData, TreeNode node, int initialValueIndex = 0)
        {
            if (fieldData == null || node == null)
            {
                throw new ArgumentNullException($"{nameof(fieldData)}|{nameof(node)}");
            }
            var nodeContent = node.Content;
            if (nodeContent.NodeLabel == "Terminal")
            {
                return;
            }

            List<ConfigExpression> valuesCollection = fieldData.Expressions;
            if (node.Children.Count == 0 && valuesCollection.Count < initialValueIndex + 1)
            {
                AddSearchValuesToChildlessNode(node, initialValueIndex - 1, valuesCollection);
            }

            if (valuesCollection.Count < initialValueIndex + 1)
            {
                return;
            }

            string fieldName = nodeContent.Name;
            if (nodeContent.NodeLabel == "Line" || nodeContent.NodeLabel == "Field")
            {
                for (int i = 0; i < node.Children.Count; i++)
                {
                    TreeNode child = node.Children[i];
                    AddSearchValuesToSingleNode(fieldName, child, valuesCollection, initialValueIndex);
                }
            }
            else
            {
                AddSearchValuesToSingleNode(fieldName, node, valuesCollection, initialValueIndex);
            }
        }
        #endregion

        #region Private static methods
        private static TreeNode AddFieldNode(TreeNode rootNode, ConfigField fieldData)
        {
            var paragraphCollection = new List<long>() { 0 };

            TreeNodeContent content = new TreeNodeContent()
            {
                Name = fieldData.Name,
                RegExPattern = fieldData.RegExPattern,
                NodeLabel = "Field",
                ValueType = fieldData.ValueType,
                CheckValue = fieldData.ExpectedName,
            };
            content.Lines.Add(paragraphCollection[0]);

            TreeNode node = new TreeNode(content);

            for (int i = 0; i < paragraphCollection.Count; i++)
            {
                TreeNodeContent childContent = new TreeNodeContent(content)
                {
                    NodeLabel = "Line"
                };
                childContent.Lines.Add(paragraphCollection[i]);
                var childNode = new TreeNode(childContent);
                node.AddChild(childNode);
            }
            rootNode.AddChild(node);
            return node;
        }
        private static void AddSearchValuesToChildlessNode(TreeNode node, int initialValueIndex, List<ConfigExpression> valuesCollection)
        {
            ConfigExpression singleValueDefinition = valuesCollection[initialValueIndex];
            TreeNodeContent content = new TreeNodeContent()
            {
                Name = node.Content.Name,
                NodeLabel = $"Search {initialValueIndex}",
                RegExPattern = singleValueDefinition.RegExPattern,
                HorizontalParagraph = ((TreeNodeContent)node.Content).HorizontalParagraph,
                ValueType = node.Content.ValueType
            };
            if (content.ValueType.Contains("Table"))
            {
                content.FirstSearchParameter = singleValueDefinition.SearchParameters["row"];
                content.SecondSearchParameter = singleValueDefinition.SearchParameters["column"];
            }
            else
            {
                content.FirstSearchParameter = singleValueDefinition.SearchParameters["line_offset"];
                content.SecondSearchParameter = singleValueDefinition.SearchParameters["horizontal_status"];
            }

            if (initialValueIndex + 1 == valuesCollection.Count)
            {
                content.NodeLabel = "Terminal";
            }
            content.Lines.Add(node.Content.Lines[0]);
            TreeNode newNode = new TreeNode(content);
            node.AddChild(newNode);
        }
        private static void AddSearchValuesToSingleNode(string fieldName, TreeNode node, List<ConfigExpression> valuesCollection, int initialValueIndex)
        {
            TreeNode singleParagraphNode = node;
            for (int valueIndex = initialValueIndex; valueIndex < valuesCollection.Count; valueIndex++)
            {
                ConfigExpression singleValueDefinition = valuesCollection[valueIndex];
                TreeNodeContent content = new TreeNodeContent()
                {
                    Name = fieldName,
                    NodeLabel = $"Search {valueIndex}",
                    RegExPattern = singleValueDefinition.RegExPattern,
                    HorizontalParagraph = ((TreeNodeContent)singleParagraphNode.Content).HorizontalParagraph,
                    ValueType = singleParagraphNode.Content.ValueType,
                };
                long offsetLine = singleParagraphNode.Content.Lines[0];

                if (content.ValueType.Contains("Table"))
                {
                    content.FirstSearchParameter = singleValueDefinition.SearchParameters["row"];
                    content.SecondSearchParameter = singleValueDefinition.SearchParameters["column"];
                }
                else
                {
                    content.FirstSearchParameter = singleValueDefinition.SearchParameters["line_offset"];
                    content.SecondSearchParameter = singleValueDefinition.SearchParameters["horizontal_status"];
                }

                if (valueIndex + 1 == valuesCollection.Count)
                {
                    content.NodeLabel = "Terminal";
                }
                content.Lines.Add(offsetLine);

                var newNode = new TreeNode(content);
                singleParagraphNode = singleParagraphNode.AddChild(newNode);
            }
        }
        #endregion
    }
}