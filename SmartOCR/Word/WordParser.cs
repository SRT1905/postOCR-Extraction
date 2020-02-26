using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    public class WordParser // TODO: add summary.
    {
        #region Private constants        
        private const long similaritySearchThreshold = 5;
        #endregion
        
        #region Fields
        private readonly ConfigData configData;
        private readonly SortedDictionary<long, List<ParagraphContainer>> lineMapping;
        private readonly List<WordTable> tables;
        private readonly SearchTree treeStructure;
        #endregion

        #region Constructors
        public WordParser(WordReader reader, ConfigData configData)
        {
            if (reader == null)
            {
                throw new ArgumentNullException(nameof(reader));
            }
            lineMapping = reader.LineMapping;
            tables = reader.TableCollection;
            treeStructure = new SearchTree(configData);
            this.configData = configData;
        }
        #endregion

        #region Public methods
        public Dictionary<string, string> ParseDocument()
        {
            treeStructure.PopulateTree();
            ProcessDocument();
            return treeStructure.GetValuesFromTree();
        }
        #endregion

        #region Private methods
        private void AddOffsetNode(TreeNode node, long searchLevel, long offsetIndex,
                                   string foundValue, decimal position, bool addToParent)
        {
            TreeNodeContent nodeContent = node.Content;
            if (nodeContent.Lines.Count(item => item == offsetIndex) >= 2)
            {
                return;
            }
            nodeContent.Lines.Add(offsetIndex);
            string nodeLabel;
            string pattern;
            decimal horizontalPosition;
            if (addToParent)
            {
                var firstChildContent = node.Children.First().Content;
                nodeLabel = firstChildContent.NodeLabel;
                pattern = firstChildContent.RegExPattern;
                horizontalPosition = position;
            }
            else
            {
                nodeLabel = node.Content.NodeLabel;
                pattern = node.Content.RegExPattern;
                horizontalPosition = position;
            }
            TreeNode childNode = node.AddChild(foundLine: offsetIndex,
                                               pattern: pattern,
                                               nodeLabel: nodeLabel,
                                               horizontalParagraph: horizontalPosition);
            childNode.Content.FoundValue = foundValue;
            SearchTree.AddSearchValues(configData[nodeContent.Name], childNode, (int)searchLevel);
        }
        private void GetDataFromUndefinedNode(TreeNode fieldNode)
        {
            Regex regexObject = Utilities.CreateRegexpObject(fieldNode.Content.RegExPattern);
            Dictionary<string, SimilarityDescription> collectedData = new Dictionary<string, SimilarityDescription>();

            var keys = lineMapping.Keys.ToList();
            for (int keyIndex = 0; keyIndex < keys.Count; keyIndex++)
            {
                long line = keys[keyIndex];
                for (int containerIndex = 0; containerIndex < lineMapping[line].Count; containerIndex++)
                {
                    ParagraphContainer container = lineMapping[line][containerIndex];
                    if (regexObject.IsMatch(container.Text))
                    {
                        var matchedDataCollection = GetMatchesFromParagraph(container.Text,
                                                                              regexObject,
                                                                              fieldNode.Content.CheckValue);

                        for (int i = 0; i < matchedDataCollection.Count; i++)
                        {
                            collectedData.Add($"{line}|{containerIndex}|{container.HorizontalLocation}|{i}", matchedDataCollection[i]);
                        }
                    }
                }
            }

            if (collectedData.Count != 0)
            {
                UpdateFieldNode(collectedData, fieldNode);
            }
        }
        private Dictionary<string, string> GetOffsetLines(long lineNumber, TreeNodeContent content)
        {
            Regex regexObject = Utilities.CreateRegexpObject(content.RegExPattern);
            var foundValuesCollection = new Dictionary<string, string>();
            List<long> keys = lineMapping.Keys.ToList();
            int lineIndex = keys.IndexOf(lineNumber);
            for (int searchOffset = 1; searchOffset <= similaritySearchThreshold; searchOffset++)
            {
                List<int> offsetIndexes = new List<int>() { lineIndex + searchOffset, lineIndex - searchOffset };
                foreach (int offsetIndex in offsetIndexes)
                {
                    if (offsetIndex >= 0 && offsetIndex < lineMapping.Count)
                    {
                        long line = keys[offsetIndex];
                        var lineChecker = new LineContentChecker(lineMapping[line]);
                        if (lineChecker.CheckLineContents(regexObject, content.CheckValue))
                        {
                            foundValuesCollection.Add(
                                string.Join("|",
                                            line,
                                            lineChecker.ParagraphHorizontalLocation),
                                lineChecker.JoinedMatches);
                        }
                    }
                }
            }
            return foundValuesCollection;
        }
        private void OffsetSearch(long lineNumber, TreeNode lineNode,
                                  long searchLevel, bool addToParent = false)
        {
            TreeNodeContent lineNodeContent = (TreeNodeContent)lineNode.Content;
            var lineNumbers = GetOffsetLines(lineNumber, lineNodeContent);

            var keys = lineNumbers.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                string[] splittedKey = key.Split('|');
                long offsetIndex = long.Parse(splittedKey[0],
                                              NumberStyles.Any,
                                              NumberFormatInfo.CurrentInfo);
                decimal horizontalPosition = decimal.Parse(splittedKey[1],
                                                           NumberStyles.Any,
                                                           NumberFormatInfo.CurrentInfo);
                if (addToParent)
                {
                    var parent = lineNode.Parent;
                    AddOffsetNode(parent, searchLevel,
                                  offsetIndex, lineNumbers[key],
                                  horizontalPosition, addToParent);
                }
                else
                {
                    AddOffsetNode(lineNode, searchLevel,
                                  offsetIndex, lineNumbers[key],
                                  horizontalPosition, addToParent);
                }
            }
        }
        private void ProcessDocument()
        {
            for (int fieldIndex = 0; fieldIndex < treeStructure.Children.Count; fieldIndex++)
            {
                TreeNode fieldNode = treeStructure.Children[fieldIndex];
                TreeNodeContent nodeContent = fieldNode.Content;
                if (!nodeContent.ValueType.Contains("Table"))
                {
                    if (nodeContent.Lines[0] == 0)
                    {
                        GetDataFromUndefinedNode(fieldNode);
                    }
                    if (nodeContent.Lines[0] != 0)
                    {
                        for (int i = 0; i < fieldNode.Children.Count; i++)
                        {
                            TreeNode lineNode = fieldNode.Children[i];
                            ProcessLineNode(lineNode);
                        }
                    }
                }
                else
                {
                    TableNodeProcessor processor = new TableNodeProcessor(tables, fieldNode);
                    processor.Process();
                }
            }
        }
        private void ProcessLineNode(TreeNode lineNode, long searchLevel = 0)
        {
            TreeNodeContent lineNodeContent = (TreeNodeContent)lineNode.Content;
            if (lineNodeContent.NodeLabel == "Terminal")
            {
                ProcessValue(lineNode, searchLevel);
                return;
            }

            int lineIndex = 0;
            while (lineIndex < lineNodeContent.Lines.Count)
            {
                long lineNumber = lineNodeContent.Lines[lineIndex];

                bool checkStatus = lineMapping.ContainsKey(lineNumber)
                    ? TryMatchLineData(lineNodeContent, lineNumber)
                    : false;
                if (checkStatus)
                {
                    SetOffsetChildrenLines(lineNode, lineNumber);
                    ProcessLineNodeChildren(lineNode, searchLevel);
                }
                else
                {
                    OffsetSearch(lineNumber, lineNode, searchLevel, true);
                }
                lineIndex++;
            }
        }
        private void ProcessLineNodeChildren(TreeNode lineNode, long searchLevel)
        {
            int childIndex = 0;
            while (childIndex < lineNode.Children.Count)
            {
                TreeNode childNode = lineNode.Children[childIndex];
                ProcessLineNode(childNode, searchLevel + 1);
                childIndex++;
            }
        }
        private void ProcessValue(TreeNode node, long searchLevel)
        {
            TreeNodeContent nodeContent = node.Content;
            for (int i = 0; i < nodeContent.Lines.Count; i++)
            {
                long lineNumber = nodeContent.Lines[i];
                if (!lineMapping.ContainsKey(lineNumber))
                {
                    OffsetSearch(lineNumber, node, searchLevel, true);
                }
                else
                {
                    List<ParagraphContainer> paragraphCollection = lineMapping[lineNumber];
                    int startIndex = 0;
                    int finishIndex = paragraphCollection.Count - 1;
                    switch (nodeContent.SecondSearchParameter)
                    {
                        case 1:
                            startIndex = GetParagraphByLocation(paragraphCollection,
                                                                 nodeContent.HorizontalParagraph,
                                                                 returnNextLargest: true);
                            finishIndex = paragraphCollection.Count - 1;
                            break;
                        case -1:
                            startIndex = 0;
                            finishIndex = GetParagraphByLocation(paragraphCollection,
                                                                  nodeContent.HorizontalParagraph,
                                                                  returnNextLargest: false);
                            break;
                        default:
                            break;
                    }

                    for (int paragraphIndex = startIndex; paragraphIndex <= finishIndex; paragraphIndex++)
                    {
                        string paragraphText = paragraphCollection[paragraphIndex].Text;
                        Regex regexObject = Utilities.CreateRegexpObject(nodeContent.RegExPattern);

                        MatchProcessor matchProcessor = new MatchProcessor(paragraphText,
                                                                            regexObject,
                                                                            nodeContent.ValueType);
                        if (!string.IsNullOrEmpty(matchProcessor.Result))
                        {
                            nodeContent.FoundValue = matchProcessor.Result;
                            nodeContent.Status = true;
                            PropagateStatusInTree(true, node);
                            return;
                        }
                    }

                    OffsetSearch(lineNumber, node, searchLevel, true);
                }
            }
        }
        private bool TryMatchLineData(TreeNodeContent lineNodeContent, long lineNumber)
        {
            decimal paragraphHorizontalLocation = lineNodeContent.HorizontalParagraph;
            Regex regexObject = Utilities.CreateRegexpObject(lineNodeContent.RegExPattern);
            var lineChecker = new LineContentChecker(lineMapping[lineNumber],
                                                     paragraphHorizontalLocation,
                                                     lineNodeContent.SecondSearchParameter);
            bool checkStatus = lineChecker.CheckLineContents(regexObject, lineNodeContent.CheckValue);
            if (checkStatus)
            {
                lineNodeContent.HorizontalParagraph = lineChecker.ParagraphHorizontalLocation;
            }
            else
            {
                lineNodeContent.HorizontalParagraph = paragraphHorizontalLocation;
            }
            lineNodeContent.FoundValue = lineChecker.JoinedMatches;
            return checkStatus;
        }
        private void SetOffsetChildrenLines(TreeNode node, long line)
        {
            TreeNodeContent nodeContent = (TreeNodeContent)node.Content;

            for (int i = 0; i < node.Children.Count; i++)
            {
                TreeNode child = node.Children[i];
                TreeNodeContent childContent = (TreeNodeContent)child.Content;
                childContent.HorizontalParagraph = nodeContent.HorizontalParagraph;
                List<long> keys = lineMapping.Keys.ToList();
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
            SearchTree.AddSearchValues(configData[fieldNode.Content.Name], fieldNode);
        }
        #endregion

        #region Private static methods
        private static void AddChildrenToFieldNode(TreeNode fieldNode,
                                                   Dictionary<string, SimilarityDescription> collectedData,
                                                   double maxSimilarity)
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
            long line = long.Parse(key.Split('|')[0],
                                   NumberStyles.Any,
                                   NumberFormatInfo.CurrentInfo);
            decimal horizontalLocation = decimal.Parse(key.Split('|')[2],
                                                        NumberStyles.Any,
                                                        NumberFormatInfo.CurrentInfo);
            if (!content.Lines.Contains(line))
            {
                content.Lines.Add(line);
                fieldNode.AddChild(foundLine: line, pattern: content.RegExPattern,
                                    newValue: content.CheckValue, nodeLabel: "Line",
                                    horizontalParagraph: horizontalLocation);
            }
        }
        private static List<SimilarityDescription> GetMatchesFromParagraph(string textToCheck,
                                                                           Regex regexObject,
                                                                           string checkValue)
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
                        SimilarityDescription description = new SimilarityDescription(groupItem.Value,
                                                                                      checkValue);
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
        private static int GetParagraphByLocation(List<ParagraphContainer> paragraphCollection,
                                                  decimal position,
                                                  bool returnNextLargest)
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
        #endregion
    }
}