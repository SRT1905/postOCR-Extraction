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
            PerformNullCheck(fieldData, node);
            if (node.Content.NodeLabel == "Terminal")
            {
                return;
            }

            if (fieldData.Expressions.Count < initialValueIndex + 1)
            {
                if (node.Children.Count != 0)
                {
                    return;
                }

                AddSearchValuesToChildlessNode(node, fieldData.Expressions, initialValueIndex - 1);
            }

            AddSearchValuesToValidNode(fieldData, node, initialValueIndex);
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
                this.GetValuesForSingleField(finalValues, field);
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
                AddSearchValues(field, AddFieldNode(root, field));
            }

            this.treeStructure = root;
        }

        private static void AddSearchValuesToValidNode(ConfigField fieldData, TreeNode node, int initialValueIndex)
        {
            if (node.Content.NodeLabel == "Line" || node.Content.NodeLabel == "Field")
            {
                AddSearchValuesToHighLevelNode(node, fieldData.Expressions, initialValueIndex);
            }
            else
            {
                AddSearchValuesToSingleNode(node, fieldData.Expressions, initialValueIndex);
            }
        }

        private static void PerformNullCheck(ConfigField fieldData, TreeNode node)
        {
            if (fieldData == null || node == null)
            {
                throw new ArgumentNullException($"{nameof(fieldData)}|{nameof(node)}");
            }
        }

        private static void AddSearchValuesToHighLevelNode(TreeNode node, List<ConfigExpression> valuesCollection, int initialValueIndex)
        {
            for (int i = 0; i < node.Children.Count; i++)
            {
                AddSearchValuesToSingleNode(node.Children[i], valuesCollection, initialValueIndex);
            }
        }

        private static TreeNode AddFieldNode(TreeNode rootNode, ConfigField fieldData)
        {
            TreeNode node = InitializeNodeFromConfigField(fieldData);
            node.AddChild(
                new TreeNode(
                    new TreeNodeContentBuilder(node.Content).SetNodeLabel("Line")
                                                            .AddLine(0)
                                                            .Build()));
            rootNode.AddChild(node);
            return node;
        }

        private static TreeNode InitializeNodeFromConfigField(ConfigField fieldData)
        {
            TreeNodeContent content = new TreeNodeContentBuilder().SetName(fieldData.Name)
                                                                  .SetTextExpression(fieldData.TextExpression)
                                                                  .SetNodeLabel("Field")
                                                                  .SetValueType(fieldData.ValueType)
                                                                  .SetCheckValue(fieldData.ExpectedName)
                                                                  .SetSoundexUsageStatus(fieldData.UseSoundex)
                                                                  .AddLine(0)
                                                                  .Build();
            return new TreeNode(content);
        }

        private static void AddSearchValuesToChildlessNode(TreeNode node, List<ConfigExpression> valuesCollection, int initialValueIndex)
        {
            node.AddChild(
                new TreeNode(
                    InitializeContentToAddSearchValues(node, valuesCollection, initialValueIndex).Build()));
        }

        private static void AddSearchValuesToSingleNode(TreeNode node, List<ConfigExpression> valuesCollection, int initialValueIndex)
        {
            TreeNode singleParagraphNode = node;
            for (int valueIndex = initialValueIndex; valueIndex < valuesCollection.Count; valueIndex++)
            {
                singleParagraphNode = singleParagraphNode.AddChild(
                    new TreeNode(
                        InitializeContentToAddSearchValues(singleParagraphNode, valuesCollection, valueIndex).Build()));
            }
        }

        private static TreeNodeContentBuilder InitializeContentToAddSearchValues(
            TreeNode node,
            List<ConfigExpression> valuesCollection,
            int initialValueIndex)
        {
            TreeNodeContentBuilder contentBuilder = new TreeNodeContentBuilder().SetName(node.Content.Name)
                                                                                .SetNodeLabel(initialValueIndex + 1 == valuesCollection.Count
                                                                                                ? "Terminal"
                                                                                                : $"Search {initialValueIndex}")
                                                                                .SetTextExpression(valuesCollection[initialValueIndex].RegExPattern)
                                                                                .SetHorizontalParagraph(node.Content.HorizontalParagraph)
                                                                                .SetValueType(node.Content.ValueType)
                                                                                .SetSoundexUsageStatus(node.Content.UseSoundex)
                                                                                .AddLine(node.Content.Lines[0]);
            contentBuilder = DefineNumericSearchParameters(valuesCollection[initialValueIndex].SearchParameters, node.Content.ValueType, contentBuilder);
            return contentBuilder;
        }

        private static TreeNodeContentBuilder DefineNumericSearchParameters(Dictionary<string, int> searchParameters, string valueType, TreeNodeContentBuilder contentBuilder)
        {
            return valueType.Contains("Table")
                ? contentBuilder.SetFirstSearchParameter(searchParameters["row"])
                                .SetSecondSearchParameter(searchParameters["column"])
                : contentBuilder.SetFirstSearchParameter(searchParameters["line_offset"])
                                .SetSecondSearchParameter(searchParameters["horizontal_status"]);
        }

        private void GetValuesForSingleField(Dictionary<string, string> finalValues, ConfigField field)
        {
            finalValues.Add(field.Name, string.Join("|", this.GetDataFromChildren(field.Name, this.GetChildrenByFieldName(field.Name))));
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

        private HashSet<string> GetDataFromChildren(string fieldName, List<string> childrenCollection)
        {
            return childrenCollection.Count != 0
                ? new HashSet<string>(childrenCollection)
                : new HashSet<string>(this.GetDataFromPreTerminalNodes(fieldName));
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