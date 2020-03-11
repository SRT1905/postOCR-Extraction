namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to find data, specified by configuration fields, in Word document.
    /// </summary>
    public class WordParser // TODO: add debug on node processing
    {
        private ConfigData configData;
        private LineMapping lineMapping;
        private List<WordTable> tables;
        private SearchTree treeStructure;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordParser"/> class.
        /// </summary>
        /// <param name="reader">Contents of single Word documents.</param>
        /// <param name="configData">A collection of configuration fields.</param>
        public WordParser(WordReader reader, ConfigData configData)
        {
            if (reader == null)
            {
                throw new ArgumentNullException(nameof(reader));
            }

            this.InitializeFields(reader, configData);
        }

        /// <summary>
        /// Performs search of data in document. Returns mapping between field name and found value.
        /// </summary>
        /// <returns>A mapping between field name and found value.</returns>
        public Dictionary<string, string> ParseDocument()
        {
            Utilities.Debug($"Parsing document.", 1);
            this.treeStructure.PopulateTree();
            this.ProcessDocument();
            return this.treeStructure.GetValuesFromTree();
        }

        private static int GetParagraphByLocation(List<ParagraphContainer> paragraphCollection, decimal position, bool returnNextLargest)
        {
            int location = paragraphCollection.Select(item => item.HorizontalLocation).ToList().BinarySearch(position);
            return GetValidatedParagraphLocation(
                paragraphCollection,
                ValidateNegativeParagraphLocation(location),
                returnNextLargest);
        }

        private static int ValidateNegativeParagraphLocation(int location)
        {
            return location < 0
                ? ~location
                : location;
        }

        private static int GetValidatedParagraphLocation(List<ParagraphContainer> paragraphCollection, int location, bool returnNextLargest)
        {
            return !returnNextLargest || location == paragraphCollection.Count
                ? --location
                : location;
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

        private void InitializeFields(WordReader reader, ConfigData configData)
        {
            this.lineMapping = reader.Mapping;
            this.tables = reader.TableCollection;
            this.treeStructure = new SearchTree(configData);
            this.configData = configData;
        }

        private void ProcessDocument()
        {
            for (int fieldIndex = 0; fieldIndex < this.treeStructure.Children.Count; fieldIndex++)
            {
                TreeNode fieldNode = this.treeStructure.Children[fieldIndex];
                if (!fieldNode.Content.ValueType.Contains("Table"))
                {
                    this.InitializeLineNodeAndGetData(fieldNode);
                }
                else
                {
                    this.InitializeTableNodeAndGetData(fieldNode);
                }
            }
        }

        private void InitializeTableNodeAndGetData(TreeNode fieldNode)
        {
            Utilities.Debug($"Performing search for table node '{fieldNode.Content.Name}' data.", 2);
            TableNodeProcessor processor = new TableNodeProcessor(this.tables, fieldNode);
            processor.Process();
        }

        private void InitializeLineNodeAndGetData(TreeNode fieldNode)
        {
            if (fieldNode.Content.Lines[0] == 0)
            {
                Utilities.Debug($"Initializing field node '{fieldNode.Content.Name}' data.", 2);
                fieldNode = new UndefinedNodeProcessor(fieldNode, this.lineMapping, this.configData).GetProcessedNode();
            }

            if (fieldNode.Content.Lines[0] != 0)
            {
                Utilities.Debug($"Performing search for field node '{fieldNode.Content.Name}' data.", 2);
                for (int i = 0; i < fieldNode.Children.Count; i++)
                {
                    TreeNode lineNode = fieldNode.Children[i];
                    this.ProcessLineNode(lineNode);
                }
            }
        }

        private void ProcessLineNode(TreeNode lineNode, int searchLevel = 0)
        {
            TreeNodeContent lineNodeContent = (TreeNodeContent)lineNode.Content;
            if (lineNodeContent.NodeLabel == "Terminal")
            {
                this.ProcessValue(lineNode, searchLevel);
                return;
            }

            int lineIndex = 0;
            while (lineIndex < lineNodeContent.Lines.Count)
            {
                int lineNumber = lineNodeContent.Lines[lineIndex];

                bool checkStatus = this.lineMapping.ContainsKey(lineNumber)
                    ? this.TryMatchLineData(lineNodeContent, lineNumber)
                    : false;
                if (checkStatus)
                {
                    this.SetOffsetChildrenLines(lineNode, lineNumber);
                    this.ProcessLineNodeChildren(lineNode, searchLevel);
                }
                else
                {
                    new OffsetNodeProcessor(this.configData[lineNodeContent.Name], this.lineMapping)
                        .OffsetSearch(lineNumber, lineNode, searchLevel, true);
                }

                lineIndex++;
            }
        }

        private void ProcessLineNodeChildren(TreeNode lineNode, int searchLevel)
        {
            int childIndex = 0;
            while (childIndex < lineNode.Children.Count)
            {
                TreeNode childNode = lineNode.Children[childIndex];
                this.ProcessLineNode(childNode, searchLevel + 1);
                childIndex++;
            }
        }

        private void ProcessValue(TreeNode node, int searchLevel)
        {
            TreeNodeContent nodeContent = node.Content;
            for (int i = 0; i < nodeContent.Lines.Count; i++)
            {
                int lineNumber = nodeContent.Lines[i];
                if (!this.lineMapping.ContainsKey(lineNumber))
                {
                    new OffsetNodeProcessor(this.configData[nodeContent.Name], this.lineMapping)
                        .OffsetSearch(lineNumber, node, searchLevel, true);
                    return;
                }

                List<ParagraphContainer> paragraphCollection = this.lineMapping[lineNumber];
                this.DefineSearchIndexesWhenProcessingValue(nodeContent, paragraphCollection, out int startIndex, out int finishIndex);

                for (int paragraphIndex = startIndex; paragraphIndex <= finishIndex; paragraphIndex++)
                {
                    if (this.ProcessValueInSingleParagraph(paragraphCollection[paragraphIndex].Text, node))
                    {
                        return;
                    }
                }

                new OffsetNodeProcessor(this.configData[nodeContent.Name], this.lineMapping)
                    .OffsetSearch(lineNumber, node, searchLevel, true);
            }
        }

        private void DefineSearchIndexesWhenProcessingValue(
            TreeNodeContent nodeContent,
            List<ParagraphContainer> paragraphCollection,
            out int startIndex,
            out int finishIndex)
        {
            startIndex = 0;
            finishIndex = paragraphCollection.Count - 1;
            switch (nodeContent.SecondSearchParameter)
            {
                case 1:
                    startIndex = GetParagraphByLocation(paragraphCollection, nodeContent.HorizontalParagraph, returnNextLargest: true);
                    break;
                case -1:
                    finishIndex = GetParagraphByLocation(paragraphCollection, nodeContent.HorizontalParagraph, returnNextLargest: false);
                    break;
                default:
                    break;
            }
        }

        private bool ProcessValueInSingleParagraph(string paragraphText, TreeNode node)
        {
            TreeNodeContent nodeContent = node.Content;

            MatchProcessor matchProcessor = new MatchProcessor(
                paragraphText,
                Utilities.CreateRegexpObject(nodeContent.TextExpression),
                nodeContent.ValueType);
            if (!string.IsNullOrEmpty(matchProcessor.Result))
            {
                nodeContent.FoundValue = matchProcessor.Result;
                nodeContent.Status = true;
                PropagateStatusInTree(true, node);
            }

            return nodeContent.Status;
        }

        private bool TryMatchLineData(TreeNodeContent lineNodeContent, int lineNumber)
        {
            decimal paragraphHorizontalLocation = lineNodeContent.HorizontalParagraph;
            Regex regexObject = Utilities.CreateRegexpObject(lineNodeContent.TextExpression);
            var lineChecker = new LineContentChecker(
                this.lineMapping[lineNumber], lineNodeContent.UseSoundex, paragraphHorizontalLocation, lineNodeContent.SecondSearchParameter);
            bool checkStatus = lineChecker.CheckLineContents(regexObject, lineNodeContent.CheckValue);
            lineNodeContent.HorizontalParagraph = checkStatus
                ? lineChecker.ParagraphHorizontalLocation
                : lineNodeContent.HorizontalParagraph = paragraphHorizontalLocation;
            lineNodeContent.FoundValue = lineChecker.JoinedMatches;
            return checkStatus;
        }

        private void SetOffsetChildrenLines(TreeNode node, int line)
        {
            TreeNodeContent nodeContent = node.Content;

            for (int i = 0; i < node.Children.Count; i++)
            {
                TreeNode child = node.Children[i];
                TreeNodeContent childContent = child.Content;
                childContent.HorizontalParagraph = nodeContent.HorizontalParagraph;
                List<int> keys = this.lineMapping.Keys.ToList();
                int lineIndex = keys.IndexOf(line) + childContent.FirstSearchParameter;
                if (lineIndex >= 0 && lineIndex < keys.Count)
                {
                    int offsetLine = keys[lineIndex];
                    childContent.Lines.Clear();
                    childContent.Lines.Add(offsetLine);
                }
            }
        }
    }
}