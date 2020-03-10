namespace SmartOCR
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents collection of <see cref="TreeNode"/> instances that are used to search information in document.
    /// </summary>
    public class SearchTree
    {
        private readonly ConfigData configData;
        private TreeNode treeStructure;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchTree"/> class.
        /// </summary>
        /// <param name="configData">An instance of <see cref="ConfigData"/> with information about search fields.</param>
        public SearchTree(ConfigData configData)
        {
            this.configData = configData;
        }

        /// <summary>
        /// Gets collection of field nodes.
        /// </summary>
        public List<TreeNode> Children
        {
            get
            {
                return this.treeStructure.Children;
            }
        }

        /// <summary>
        /// Adds instances of <see cref="TreeNode"/> class to <paramref name="node"/> that defined by search expressions in <paramref name="fieldData"/>, starting from <paramref name="initialValueIndex"/>.
        /// </summary>
        /// <param name="fieldData">An instance of <see cref="ConfigField"/> class containing search field definition and search expressions.</param>
        /// <param name="node">An instance of <see cref="TreeNode"/> class, to which new nodes are to be added.</param>
        /// <param name="initialValueIndex">Starting level of search expressions to add.</param>
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
            if (valuesCollection.Count < initialValueIndex + 1)
            {
                if (node.Children.Count != 0)
                {
                    return;
                }

                AddSearchValuesToChildlessNode(node, initialValueIndex - 1, valuesCollection);
            }

            if (nodeContent.NodeLabel == "Line" || nodeContent.NodeLabel == "Field")
            {
                for (int i = 0; i < node.Children.Count; i++)
                {
                    TreeNode child = node.Children[i];
                    AddSearchValuesToSingleNode(nodeContent.Name, child, valuesCollection, initialValueIndex);
                }
            }
            else
            {
                AddSearchValuesToSingleNode(nodeContent.Name, node, valuesCollection, initialValueIndex);
            }
        }

        /// <summary>
        /// Loops through field nodes and collects found values.
        /// </summary>
        /// <returns>A mapping between field names and found values.</returns>
        public Dictionary<string, string> GetValuesFromTree()
        {
            Utilities.Debug($"Collecting found data from search tree.", 2);
            var finalValues = new Dictionary<string, string>();
            foreach (ConfigField field in this.configData.Fields)
            {
                List<string> childrenCollection = this.GetChildrenByFieldName(field.Name);
                var result = new HashSet<string>();
                if (childrenCollection.Count != 0)
                {
                    result.UnionWith(childrenCollection);
                }
                else
                {
                    var preTerminalCollection = this.GetDataFromPreTerminalNodes(field.Name);
                    result.UnionWith(preTerminalCollection);
                }

                finalValues.Add(field.Name, string.Join("|", result));
            }

            return finalValues;
        }

        /// <summary>
        /// Initializes a search tree with data from instance of <see cref="ConfigData"/>, provided in constructor.
        /// </summary>
        public void PopulateTree()
        {
            Utilities.Debug($"Populating search tree with config data.", 2);
            TreeNode root = TreeNode.CreateRoot();
            foreach (ConfigField field in this.configData.Fields)
            {
                TreeNode fieldNode = AddFieldNode(root, field);
                AddSearchValues(field, fieldNode);
            }

            this.treeStructure = root;
        }

        private static TreeNode AddFieldNode(TreeNode rootNode, ConfigField fieldData)
        {
            var paragraphCollection = new List<int>() { 0 };

            TreeNodeContent content = new TreeNodeContent()
            {
                Name = fieldData.Name,
                TextExpression = fieldData.TextExpression,
                NodeLabel = "Field",
                ValueType = fieldData.ValueType,
                CheckValue = fieldData.ExpectedName,
                UseSoundex = fieldData.UseSoundex,
            };
            content.Lines.Add(paragraphCollection[0]);

            TreeNode node = new TreeNode(content);

            for (int i = 0; i < paragraphCollection.Count; i++)
            {
                TreeNodeContent childContent = new TreeNodeContent(content)
                {
                    NodeLabel = "Line",
                };
                childContent.Lines.Add(paragraphCollection[i]);
                node.AddChild(new TreeNode(childContent));
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
                NodeLabel = initialValueIndex + 1 == valuesCollection.Count
                    ? "Terminal"
                    : $"Search {initialValueIndex}",
                TextExpression = singleValueDefinition.RegExPattern,
                HorizontalParagraph = node.Content.HorizontalParagraph,
                ValueType = node.Content.ValueType,
                UseSoundex = node.Content.UseSoundex,
            };
            DefineNumericSearchParameters(singleValueDefinition, content);

            content.Lines.Add(node.Content.Lines[0]);
            node.AddChild(new TreeNode(content));
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
                    NodeLabel = valueIndex + 1 == valuesCollection.Count
                        ? "Terminal"
                        : $"Search {valueIndex}",
                    TextExpression = singleValueDefinition.RegExPattern,
                    HorizontalParagraph = singleParagraphNode.Content.HorizontalParagraph,
                    ValueType = singleParagraphNode.Content.ValueType,
                    UseSoundex = singleParagraphNode.Content.UseSoundex,
                };

                DefineNumericSearchParameters(singleValueDefinition, content);

                content.Lines.Add(singleParagraphNode.Content.Lines[0]);
                singleParagraphNode = singleParagraphNode.AddChild(new TreeNode(content));
            }
        }

        private static void DefineNumericSearchParameters(ConfigExpression singleValueDefinition, TreeNodeContent content)
        {
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
        }

        private List<string> GetChildrenByFieldName(string fieldName)
        {
            var childrenCollection = new List<string>();
            foreach (TreeNode fieldNode in this.treeStructure.Children)
            {
                if (fieldNode.Content.Name == fieldName)
                {
                    this.GetNodeChildren(fieldNode, childrenCollection);
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
                this.GetDataFromNode(child, foundData);
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
                { false, new HashSet<string>() },
            };
            foreach (TreeNode node in this.treeStructure.Children)
            {
                if (node.Content.Name == fieldName)
                {
                    this.GetDataFromNode(node, foundData);
                    return foundData[foundData[true].Count != 0];
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
                this.GetNodeChildren(child, childrenCollection);
            }
        }
    }
}