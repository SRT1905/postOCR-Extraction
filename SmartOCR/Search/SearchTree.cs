namespace SmartOCR.Search
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using SmartOCR.Config;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

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
        public List<TreeNode> Children => this.treeStructure.Children;

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

            ValidateParametersAndAddValues(fieldData, node, initialValueIndex);
        }

        /// <summary>
        /// Loops through field nodes and collects found values.
        /// </summary>
        /// <returns>A mapping between field names and found values.</returns>
        public Dictionary<string, string> GetValuesFromTree()
        {
            Utilities.Debug("Collecting found data from search tree.", 2);
            var finalValues = new Dictionary<string, string>();
            foreach (var field in this.configData.Fields)
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
            Utilities.Debug("Populating search tree with config data.", 2);
            this.treeStructure = this.PopulateRoot();
        }

        private static void ValidateParametersAndAddValues(ConfigField fieldData, TreeNode node, int initialValueIndex)
        {
            if (fieldData.Expressions.Count < initialValueIndex + 1 && node.Children.Count != 0)
            {
                return;
            }

            AddValuesToValidNode(fieldData, node, initialValueIndex);
        }

        private static void AddValuesToValidNode(ConfigField fieldData, TreeNode node, int initialValueIndex)
        {
            if (fieldData.Expressions.Count < initialValueIndex + 1)
            {
                AddSearchValuesToChildlessNode(node, fieldData.Expressions, initialValueIndex - 1);
            }
            else
            {
                AddSearchValuesToDefaultNode(fieldData, node, initialValueIndex);
            }
        }

        private static void AddSearchValuesToDefaultNode(ConfigField fieldData, TreeNode node, int initialValueIndex)
        {
            if (new List<string> { "Line", "Field" }.Contains(node.Content.NodeLabel))
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
            foreach (var childNode in node.Children)
            {
                AddSearchValuesToSingleNode(childNode, valuesCollection, initialValueIndex);
            }
        }

        private static TreeNode AddFieldNode(TreeNode rootNode, ConfigField fieldData)
        {
            var node = InitializeNodeFromConfigField(fieldData);
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
            return new TreeNode(new TreeNodeContentBuilder().SetName(fieldData.Name)
                                                            .SetTextExpression(fieldData.TextExpression)
                                                            .SetNodeLabel("Field")
                                                            .SetValueType(fieldData.ValueType)
                                                            .SetCheckValue(fieldData.ExpectedName)
                                                            .SetSoundexUsageStatus(fieldData.UseSoundex)
                                                            .AddLine(0)
                                                            .SetGridCoordinates(fieldData.GridCoordinates)
                                                            .Build());
        }

        private static void AddSearchValuesToChildlessNode(TreeNode node, List<ConfigExpression> valuesCollection, int initialValueIndex)
        {
            node.AddChild(
                new TreeNode(
                    InitializeContentToAddSearchValues(node, valuesCollection, initialValueIndex).Build()));
        }

        private static void AddSearchValuesToSingleNode(TreeNode node, List<ConfigExpression> valuesCollection, int initialValueIndex)
        {
            var singleParagraphNode = node;
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
            TreeNodeContentBuilder contentBuilder = new TreeNodeContentBuilder(node.Content) // .SetName(node.Content.Name)
                                                                                .SetNodeLabel(initialValueIndex + 1 == valuesCollection.Count
                                                                                                ? "Terminal"
                                                                                                : $"Search {initialValueIndex}")
                                                                                .SetTextExpression(valuesCollection[initialValueIndex].RegExPattern)
                                                                                .AddLine(node.Content.Lines[0]);
            return DefineNumericSearchParameters(valuesCollection[initialValueIndex].SearchParameters, node.Content.ValueType, contentBuilder);
        }

        private static TreeNodeContentBuilder DefineNumericSearchParameters(Dictionary<string, int> searchParameters, string valueType, TreeNodeContentBuilder contentBuilder)
        {
            return valueType.Contains("Table")
                ? contentBuilder.SetFirstSearchParameter(searchParameters["row"])
                                .SetSecondSearchParameter(searchParameters["column"])
                : contentBuilder.SetFirstSearchParameter(searchParameters["line_offset"])
                                .SetSecondSearchParameter(searchParameters["horizontal_status"]);
        }

        private static void GetDataFromNodeWithNonEmptyValue(TreeNode node, Dictionary<bool, HashSet<string>> foundData)
        {
            if (node.Content.Status || node.Parent.Content.NodeLabel != "Field")
            {
                foundData[node.Content.Status].Add(node.Content.FoundValue);
            }
        }

        private TreeNode PopulateRoot()
        {
            TreeNode root = TreeNode.CreateRoot();

            foreach (ConfigField field in this.configData.Fields)
            {
                AddSearchValues(field, AddFieldNode(root, field));
            }

            return root;
        }

        private void GetValuesForSingleField(Dictionary<string, string> finalValues, ConfigField field)
        {
            finalValues.Add(field.Name, string.Join("|", this.GetDataFromChildren(field.Name, this.GetChildrenByFieldName(field.Name))));
        }

        private List<string> GetChildrenByFieldName(string fieldName)
        {
            var childrenCollection = new List<string>();
            foreach (var fieldNode in this.treeStructure.Children.Where(fieldNode => fieldNode.Content.Name == fieldName))
            {
                this.GetNodeChildren(fieldNode, childrenCollection);
                break;
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

            foreach (var childNode in node.Children)
            {
                this.GetDataFromNode(childNode, foundData);
            }

            if (!string.IsNullOrEmpty(node.Content.FoundValue))
            {
                GetDataFromNodeWithNonEmptyValue(node, foundData);
            }
        }

        private HashSet<string> GetDataFromPreTerminalNodes(string fieldName)
        {
            var foundData = new Dictionary<bool, HashSet<string>>()
            {
                { true, new HashSet<string>() },
                { false, new HashSet<string>() },
            };
            foreach (var node in this.treeStructure.Children.Where(node => node.Content.Name == fieldName))
            {
                this.GetDataFromNode(node, foundData);
                return foundData[foundData[true].Count != 0];
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

            foreach (var childNode in node.Children)
            {
                this.GetNodeChildren(childNode, childrenCollection);
            }
        }
    }
}