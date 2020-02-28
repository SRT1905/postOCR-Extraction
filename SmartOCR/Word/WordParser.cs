namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to find data, specified by configuration fields, in Word document.
    /// </summary>
    public class WordParser
    {
        private const long SimilaritySearchThreshold = 5;

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
            long line = long.Parse(key.Split('|')[0], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
            decimal horizontalLocation = decimal.Parse(key.Split('|')[2], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
            if (!content.Lines.Contains(line))
            {
                content.Lines.Add(line);
                var builder = new TreeNodeContentBuilder(content);
                builder.SetNodeLabel("Line");
                builder.SetHorizontalParagraph(horizontalLocation);
                fieldNode.AddChild(builder.Build());
            }
        }

        private static List<SimilarityDescription> GetMatchesFromParagraph(string textToCheck, Regex regexObject, string checkValue)
        {
            MatchCollection matches = regexObject.Matches(textToCheck);
            var foundValues = new List<SimilarityDescription>();
            for (int i = 0; i < matches.Count; i++)
            {
                Match singleMatch = matches[i];
                if (singleMatch.Groups.Count > 1)
                {
                    for (int groupIndex = 1; groupIndex < singleMatch.Groups.Count; groupIndex++)
                    {
                        Group groupItem = singleMatch.Groups[groupIndex];
                        SimilarityDescription description = new SimilarityDescription(groupItem.Value, checkValue);
                        if (description.CheckStringSimilarity())
                        {
                            foundValues.Add(description);
                        }
                    }
                }
                else
                {
                    SimilarityDescription description = new SimilarityDescription(singleMatch.Value, checkValue);
                    if (description.CheckStringSimilarity())
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

        private void AddOffsetNode(TreeNode node, int searchLevel, long offsetIndex, string foundValue, decimal position, bool addToParent)
        {
            if (node.Content.Lines.Count(item => item == offsetIndex) >= 2)
            {
                return;
            }

            var contentBuilder = new TreeNodeContentBuilder(node.Content);
            contentBuilder.TryAddLine(offsetIndex);
            contentBuilder.SetHorizontalParagraph(position);
            contentBuilder.SetNodeLabel(node.Content.NodeLabel);
            contentBuilder.SetRegExPattern(node.Content.RegExPattern);

            if (addToParent)
            {
                contentBuilder.SetNodeLabel(node.Children[0].Content.NodeLabel);
                contentBuilder.SetRegExPattern(node.Children[0].Content.RegExPattern);
            }

            TreeNode childNode = node.AddChild(contentBuilder.Build());
            childNode.Content.FoundValue = foundValue;
            SearchTree.AddSearchValues(this.configData[childNode.Content.Name], childNode, searchLevel);
        }

        private void GetDataFromUndefinedNode(TreeNode fieldNode)
        {
            Regex regexObject = Utilities.CreateRegexpObject(fieldNode.Content.RegExPattern);
            Dictionary<string, SimilarityDescription> collectedData = new Dictionary<string, SimilarityDescription>();

            var keys = this.lineMapping.Keys.ToList();
            for (int keyIndex = 0; keyIndex < keys.Count; keyIndex++)
            {
                long line = keys[keyIndex];
                for (int containerIndex = 0; containerIndex < this.lineMapping[line].Count; containerIndex++)
                {
                    ParagraphContainer container = this.lineMapping[line][containerIndex];
                    if (regexObject.IsMatch(container.Text))
                    {
                        var matchedDataCollection = GetMatchesFromParagraph(container.Text, regexObject, fieldNode.Content.CheckValue);

                        for (int i = 0; i < matchedDataCollection.Count; i++)
                        {
                            collectedData.Add($"{line}|{containerIndex}|{container.HorizontalLocation}|{i}", matchedDataCollection[i]);
                        }
                    }
                }
            }

            if (collectedData.Count != 0)
            {
                this.UpdateFieldNode(collectedData, fieldNode);
            }
        }

        private Dictionary<string, string> GetOffsetLines(long lineNumber, TreeNodeContent content)
        {
            Regex regexObject = Utilities.CreateRegexpObject(content.RegExPattern);
            var foundValuesCollection = new Dictionary<string, string>();
            List<long> keys = this.lineMapping.Keys.ToList();
            int lineIndex = keys.IndexOf(lineNumber);
            for (int searchOffset = 1; searchOffset <= SimilaritySearchThreshold; searchOffset++)
            {
                List<int> offsetIndexes = new List<int>() { lineIndex + searchOffset, lineIndex - searchOffset };
                foreach (int offsetIndex in offsetIndexes)
                {
                    if (offsetIndex >= 0 && offsetIndex < this.lineMapping.Count)
                    {
                        long line = keys[offsetIndex];
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

        private void OffsetSearch(long lineNumber, TreeNode lineNode, int searchLevel, bool addToParent = false)
        {
            TreeNodeContent lineNodeContent = (TreeNodeContent)lineNode.Content;
            var lineNumbers = this.GetOffsetLines(lineNumber, lineNodeContent);

            var keys = lineNumbers.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                string[] splittedKey = key.Split('|');
                long offsetIndex = long.Parse(splittedKey[0], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
                decimal horizontalPosition = decimal.Parse(splittedKey[1], NumberStyles.Any, NumberFormatInfo.CurrentInfo);
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
                TreeNodeContent nodeContent = fieldNode.Content;
                if (!nodeContent.ValueType.Contains("Table"))
                {
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
                else
                {
                    Utilities.Debug($"Performing search for table node '{fieldNode.Content.Name}' data.", 2);
                    TableNodeProcessor processor = new TableNodeProcessor(this.tables, fieldNode);
                    processor.Process();
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
                long lineNumber = lineNodeContent.Lines[lineIndex];

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
                long lineNumber = nodeContent.Lines[i];
                if (!this.lineMapping.ContainsKey(lineNumber))
                {
                    this.OffsetSearch(lineNumber, node, searchLevel, true);
                }
                else
                {
                    List<ParagraphContainer> paragraphCollection = this.lineMapping[lineNumber];
                    int startIndex = 0;
                    int finishIndex = paragraphCollection.Count - 1;
                    switch (nodeContent.SecondSearchParameter)
                    {
                        case 1:
                            startIndex = GetParagraphByLocation(paragraphCollection, nodeContent.HorizontalParagraph, returnNextLargest: true);
                            finishIndex = paragraphCollection.Count - 1;
                            break;
                        case -1:
                            startIndex = 0;
                            finishIndex = GetParagraphByLocation(paragraphCollection, nodeContent.HorizontalParagraph, returnNextLargest: false);
                            break;
                        default:
                            break;
                    }

                    for (int paragraphIndex = startIndex; paragraphIndex <= finishIndex; paragraphIndex++)
                    {
                        string paragraphText = paragraphCollection[paragraphIndex].Text;
                        Regex regexObject = Utilities.CreateRegexpObject(nodeContent.RegExPattern);

                        MatchProcessor matchProcessor = new MatchProcessor(paragraphText, regexObject, nodeContent.ValueType);
                        if (!string.IsNullOrEmpty(matchProcessor.Result))
                        {
                            nodeContent.FoundValue = matchProcessor.Result;
                            nodeContent.Status = true;
                            PropagateStatusInTree(true, node);
                            return;
                        }
                    }

                    this.OffsetSearch(lineNumber, node, searchLevel, true);
                }
            }
        }

        private bool TryMatchLineData(TreeNodeContent lineNodeContent, long lineNumber)
        {
            decimal paragraphHorizontalLocation = lineNodeContent.HorizontalParagraph;
            Regex regexObject = Utilities.CreateRegexpObject(lineNodeContent.RegExPattern);
            var lineChecker = new LineContentChecker(this.lineMapping[lineNumber], paragraphHorizontalLocation, lineNodeContent.SecondSearchParameter);
            bool checkStatus = lineChecker.CheckLineContents(regexObject, lineNodeContent.CheckValue);
            lineNodeContent.HorizontalParagraph = checkStatus
                ? lineChecker.ParagraphHorizontalLocation
                : lineNodeContent.HorizontalParagraph = paragraphHorizontalLocation;
            lineNodeContent.FoundValue = lineChecker.JoinedMatches;
            return checkStatus;
        }

        private void SetOffsetChildrenLines(TreeNode node, long line)
        {
            TreeNodeContent nodeContent = node.Content;

            for (int i = 0; i < node.Children.Count; i++)
            {
                TreeNode child = node.Children[i];
                TreeNodeContent childContent = child.Content;
                childContent.HorizontalParagraph = nodeContent.HorizontalParagraph;
                List<long> keys = this.lineMapping.Keys.ToList();
                int lineIndex = keys.IndexOf(line) + childContent.FirstSearchParameter;
                if (lineIndex >= 0 && lineIndex < keys.Count)
                {
                    long offsetLine = keys[lineIndex];
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