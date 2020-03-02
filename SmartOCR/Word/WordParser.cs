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
        private const int SimilaritySearchThreshold = 5;

        private readonly ConfigData configData;
        private readonly LineMapping lineMapping;
        private readonly List<WordTable> tables;
        private readonly SearchTree treeStructure;

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

            this.lineMapping = reader.Mapping;
            this.tables = reader.TableCollection;
            this.treeStructure = new SearchTree(configData);
            this.configData = configData;
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

        private static void AddChildrenToFieldNode(TreeNode fieldNode, Dictionary<string, SimilarityDescription> collectedData, double maxSimilarity)
        {
            var keys = collectedData.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                if (maxSimilarity == collectedData[key].Ratio)
                {
                    AddSingleChildToFieldNode(fieldNode, key);
                }
            }
        }

        private static void AddSingleChildToFieldNode(TreeNode fieldNode, string key)
        {
            var content = fieldNode.Content;
            int line = int.Parse(key.Split('|')[0]);
            decimal horizontalLocation = decimal.Parse(key.Split('|')[2]);
            if (!content.Lines.Contains(line))
            {
                content.Lines.Add(line);
                var builder = new TreeNodeContentBuilder(content);
                builder.SetNodeLabel("Line");
                builder.SetHorizontalParagraph(horizontalLocation);
                fieldNode.AddChild(builder.Build());
            }
        }

        private static List<SimilarityDescription> GetMatchesFromParagraph(ParagraphContainer container, TreeNodeContent nodeContent)
        {
            MatchCollection matches = Utilities.CreateRegexpObject(nodeContent.TextExpression).Matches(container.Text);
            var foundValues = new List<SimilarityDescription>();
            for (int i = 0; i < matches.Count; i++)
            {
                int startIndex = matches[i].Groups.Count > 1
                    ? 1
                    : 0;
                for (int groupIndex = startIndex; groupIndex < matches[i].Groups.Count; groupIndex++)
                {
                    SimilarityDescription description = new SimilarityDescription(matches[i].Groups[groupIndex].Value, nodeContent.CheckValue);
                    if (description.AreStringsSimilar())
                    {
                        foundValues.Add(description);
                    }
                }
            }

            return foundValues;
        }

        private static int GetParagraphByLocation(List<ParagraphContainer> paragraphCollection, decimal position, bool returnNextLargest)
        {
            List<decimal> locations = paragraphCollection.Select(item => item.HorizontalLocation).ToList();
            int location = locations.BinarySearch(position);
            if (location < 0)
            {
                location = ~location;
            }

            if (returnNextLargest)
            {
                if (location == paragraphCollection.Count)
                {
                    return location--;
                }

                return location;
            }

            return location--;
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

        private static List<SimilarityDescription> GetSoundexMatchesFromParagraph(ParagraphContainer container, TreeNodeContent nodeContent)
        {
            List<SimilarityDescription> matches = new List<SimilarityDescription>();
            foreach (string soundexWord in container.Soundex.Split(' '))
            {
                SimilarityDescription description = new SimilarityDescription(soundexWord, nodeContent.TextExpression);
                if (description.AreStringsSimilar())
                {
                    matches.Add(description);
                }
            }

            return matches;
        }

        private void AddOffsetNode(TreeNode node, int searchLevel, int offsetIndex, string foundValue, decimal position, bool addToParent)
        {
            if (node.Content.Lines.Count(item => item == offsetIndex) >= 2)
            {
                return;
            }

            var contentBuilder = new TreeNodeContentBuilder(node.Content);
            contentBuilder.TryAddLine(offsetIndex);
            contentBuilder.SetHorizontalParagraph(position);
            contentBuilder.SetNodeLabel(node.Content.NodeLabel);
            contentBuilder.SetRegExPattern(node.Content.TextExpression);

            if (addToParent)
            {
                contentBuilder.SetNodeLabel(node.Children[0].Content.NodeLabel);
                contentBuilder.SetRegExPattern(node.Children[0].Content.TextExpression);
            }

            TreeNode childNode = node.AddChild(contentBuilder.Build());
            childNode.Content.FoundValue = foundValue;
            SearchTree.AddSearchValues(this.configData[childNode.Content.Name], childNode, searchLevel);
        }

        private void GetDataFromUndefinedNode(TreeNode fieldNode)
        {
            Dictionary<string, SimilarityDescription> collectedData = this.ProcessUndefinedNode(fieldNode);
            if (collectedData.Count != 0)
            {
                this.UpdateFieldNode(collectedData, fieldNode);
            }
        }

        private Dictionary<string, SimilarityDescription> ProcessUndefinedNode(TreeNode fieldNode)
        {
            Dictionary<string, SimilarityDescription> collectedData = new Dictionary<string, SimilarityDescription>();

            foreach (var item in this.lineMapping)
            {
                for (int containerIndex = 0; containerIndex < item.Value.Count; containerIndex++)
                {
                    ParagraphContainer container = item.Value[containerIndex];
                    List<SimilarityDescription> matches = fieldNode.Content.UseSoundex
                        ? GetSoundexMatchesFromParagraph(container, fieldNode.Content)
                        : GetMatchesFromParagraph(container, fieldNode.Content);
                    for (int matchIndex = 0; matchIndex < matches.Count; matchIndex++)
                    {
                        collectedData.Add($"{item.Key}|{containerIndex}|{container.HorizontalLocation}|{matchIndex}", matches[matchIndex]);
                    }
                }
            }

            return collectedData;
        }

        private Dictionary<string, string> GetOffsetLines(int lineNumber, TreeNodeContent content)
        {
            Regex regexObject = Utilities.CreateRegexpObject(content.TextExpression);
            var foundValuesCollection = new Dictionary<string, string>();
            List<int> keys = this.lineMapping.Keys.ToList();
            int lineIndex = keys.IndexOf(lineNumber);
            for (int searchOffset = 1; searchOffset <= SimilaritySearchThreshold; searchOffset++)
            {
                List<int> offsetIndexes = new List<int>() { lineIndex + searchOffset, lineIndex - searchOffset };
                foreach (int offsetIndex in offsetIndexes)
                {
                    if (offsetIndex >= 0 && offsetIndex < this.lineMapping.Count)
                    {
                        int line = keys[offsetIndex];
                        var lineChecker = new LineContentChecker(this.lineMapping[line]);
                        if (lineChecker.CheckLineContents(regexObject, content.CheckValue))
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

        private void OffsetSearch(int lineNumber, TreeNode lineNode, int searchLevel, bool addToParent = false)
        {
            TreeNodeContent lineNodeContent = (TreeNodeContent)lineNode.Content;
            var lineNumbers = this.GetOffsetLines(lineNumber, lineNodeContent);

            var keys = lineNumbers.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                string[] splittedKey = key.Split('|');
                int offsetIndex = int.Parse(splittedKey[0]);
                decimal horizontalPosition = decimal.Parse(splittedKey[1]);
                if (addToParent)
                {
                    var parent = lineNode.Parent;
                    this.AddOffsetNode(parent, searchLevel, offsetIndex, lineNumbers[key], horizontalPosition, addToParent);
                }
                else
                {
                    this.AddOffsetNode(lineNode, searchLevel, offsetIndex, lineNumbers[key], horizontalPosition, addToParent);
                }
            }
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
            TreeNodeContent nodeContent = fieldNode.Content;
            if (nodeContent.Lines[0] == 0)
            {
                Utilities.Debug($"Initializing field node '{fieldNode.Content.Name}' data.", 2);
                this.GetDataFromUndefinedNode(fieldNode);
            }

            if (nodeContent.Lines[0] != 0)
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
                    this.OffsetSearch(lineNumber, lineNode, searchLevel, true);
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
                    this.OffsetSearch(lineNumber, node, searchLevel, true);
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

                this.OffsetSearch(lineNumber, node, searchLevel, addToParent: true);
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
            var lineChecker = new LineContentChecker(this.lineMapping[lineNumber], paragraphHorizontalLocation, lineNodeContent.SecondSearchParameter);
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

        private void UpdateFieldNode(Dictionary<string, SimilarityDescription> collectedData, TreeNode fieldNode)
        {
            fieldNode.Children.Clear();
            if (fieldNode.Content.Lines.Count != 0)
            {
                fieldNode.Content.Lines.RemoveAt(0);
            }

            double maxSimilarity = collectedData.Values.ToList().Max(item => item.Ratio);
            AddChildrenToFieldNode(fieldNode, collectedData, maxSimilarity);
            SearchTree.AddSearchValues(this.configData[fieldNode.Content.Name], fieldNode);
        }
    }
}