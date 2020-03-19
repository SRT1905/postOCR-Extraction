namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to initialize field undefinedNode with high-level search information.
    /// </summary>
    public class UndefinedNodeProcessor
    {
        private ConfigField configField;
        private LineMapping lineMapping;
        private TreeNode node;

        /// <summary>
        /// Initializes a new instance of the <see cref="UndefinedNodeProcessor"/> class.
        /// </summary>
        /// <param name="fieldNode">An empty field undefinedNode.</param>
        /// <param name="lineMapping">A mapping between document line and paragraphs on line.</param>
        /// <param name="configField">Config field description.</param>
        public UndefinedNodeProcessor(TreeNode fieldNode, LineMapping lineMapping, ConfigField configField)
        {
            this.InitializeFields(fieldNode, lineMapping, configField);
            this.GetDataFromUndefinedNode();
        }

        /// <summary>
        /// Returns a <see cref="TreeNode"/> instance that has child line nodes.
        /// </summary>
        /// <returns>A field undefinedNode instance.</returns>
        public TreeNode GetProcessedNode()
        {
            return this.node;
        }

        private static double GetMaxSimilarityRatio(Dictionary<string, SimilarityDescription> collectedData)
        {
            return collectedData.Values.Max(item => item.Ratio);
        }

        private static string CreateSimilarityDictionaryKey(
            KeyValuePair<int, List<ParagraphContainer>> lineContents,
            ParagraphContainer container,
            int matchIndex)
        {
            return $"{lineContents.Key}|{lineContents.Value.IndexOf(container)}|{container.HorizontalLocation}|{matchIndex}";
        }

        private static string[] GetCollectedDataWithMaxSimilarity(Dictionary<string, SimilarityDescription> collectedData)
        {
            return collectedData.Where(pair => pair.Value.Ratio == GetMaxSimilarityRatio(collectedData))
                                                            .Select(pair => pair.Key)
                                                            .ToArray();
        }

        private void InitializeFields(TreeNode undefinedNode, LineMapping mapping, ConfigField configData)
        {
            this.node = undefinedNode;
            this.lineMapping = mapping;
            this.configField = configData;
        }

        private void GetDataFromUndefinedNode()
        {
            Dictionary<string, SimilarityDescription> collectedData = this.ProcessUndefinedNode();
            if (collectedData.Count != 0)
            {
                this.UpdateFieldNode(collectedData);
            }
        }

        private Dictionary<string, SimilarityDescription> ProcessUndefinedNode()
        {
            Utilities.Debug($"Gathering initial matches for field '{this.node.Content.Name}'.", 3);
            var collectedData = new Dictionary<string, SimilarityDescription>();

            foreach (var lineContents in this.lineMapping)
            {
                this.ProcessLineForUndefinedNode(collectedData, lineContents);
            }

            Utilities.Debug($"Found {collectedData.Count} initial matches for field '{this.node.Content.Name}'.", 3);
            return collectedData;
        }

        private void ProcessLineForUndefinedNode(
            Dictionary<string, SimilarityDescription> collectedData,
            KeyValuePair<int, List<ParagraphContainer>> lineContents)
        {
            foreach (var paragraph in lineContents.Value)
            {
                this.ProcessParagraphForUndefinedNode(collectedData, lineContents, paragraph);
            }
        }

        private void ProcessParagraphForUndefinedNode(
            Dictionary<string, SimilarityDescription> collectedData,
            KeyValuePair<int, List<ParagraphContainer>> lineContents,
            ParagraphContainer container)
        {
            List<SimilarityDescription> matchedDescriptions = this.GetSimilarityDescriptionCollection(container);
            foreach (SimilarityDescription description in matchedDescriptions)
            {
                Utilities.Debug($"Found initial match at line {lineContents.Key} in paragraph '{container}'.", 4);
                Utilities.Debug($"{description}.", 5);
                collectedData.Add(
                    CreateSimilarityDictionaryKey(lineContents, container, matchedDescriptions.IndexOf(description)),
                    description);
            }
        }

        private List<SimilarityDescription> GetSimilarityDescriptionCollection(ParagraphContainer container)
        {
            return this.node.Content.UseSoundex
                ? this.GetSoundexMatchesFromParagraph(container)
                : this.GetMatchesFromParagraph(container);
        }

        private List<SimilarityDescription> GetSoundexMatchesFromParagraph(ParagraphContainer container)
        {
            List<SimilarityDescription> descriptions = new List<SimilarityDescription>();
            foreach (Match item in Utilities.CreateRegexpObject(this.node.Content.TextExpression).Matches(container.Soundex))
            {
                this.GetSingleSoundexMatchFromParagraph(descriptions, item);
            }

            return descriptions;
        }

        private void GetSingleSoundexMatchFromParagraph(List<SimilarityDescription> descriptions, Match match)
        {
            SimilarityDescription description = new SimilarityDescription(match.Value, this.node.Content.TextExpression);
            if (description.AreStringsSimilar())
            {
                descriptions.Add(description);
            }
        }

        private List<SimilarityDescription> GetMatchesFromParagraph(ParagraphContainer container)
        {
            var foundValues = new List<SimilarityDescription>();
            foreach (Match match in Utilities.CreateRegexpObject(this.node.Content.TextExpression).Matches(container.Text))
            {
                this.GetSingleMatchFromParagraph(match, foundValues);
            }

            return foundValues;
        }

        private void GetSingleMatchFromParagraph(Match match, List<SimilarityDescription> foundValues)
        {
            this.GetMatchGroupsFromParagraph(match.Groups.Count > 1 ? 1 : 0, match, foundValues);
        }

        private void GetMatchGroupsFromParagraph(int groupIndex, Match match, List<SimilarityDescription> foundValues)
        {
            for (int index = groupIndex; index < match.Groups.Count; index++)
            {
                this.PerformSimilarityCheck(foundValues, match.Groups[index]);
            }
        }

        private void PerformSimilarityCheck(List<SimilarityDescription> foundValues, Group singleGroup)
        {
            var description = new SimilarityDescription(singleGroup.Value, this.node.Content.CheckValue);
            if (description.AreStringsSimilar())
            {
                foundValues.Add(description);
            }
        }

        private void UpdateFieldNode(Dictionary<string, SimilarityDescription> collectedData)
        {
            this.ClearNodeChildrenAndLines();
            this.AddChildrenToFieldNode(collectedData);
            SearchTree.AddSearchValues(this.configField, this.node);
        }

        private void ClearNodeChildrenAndLines()
        {
            this.node.Children.Clear();
            if (this.node.Content.Lines.Count != 0)
            {
                this.node.Content.Lines.RemoveAt(0);
            }
        }

        private void AddChildrenToFieldNode(Dictionary<string, SimilarityDescription> collectedData)
        {
            foreach (string key in GetCollectedDataWithMaxSimilarity(collectedData))
            {
                this.AddSingleChildToFieldNode(key);
            }
        }

        private void AddSingleChildToFieldNode(string key)
        {
            this.TryAddLineNode(Convert.ToInt32(key.Split('|')[0]), Convert.ToDecimal(key.Split('|')[2]));
        }

        private void TryAddLineNode(int line, decimal horizontalLocation)
        {
            if (!this.node.Content.Lines.Contains(line))
            {
                this.AddLineNodeToFieldNode(line, horizontalLocation);
            }
        }

        private void AddLineNodeToFieldNode(int line, decimal horizontalLocation)
        {
            this.node.Content.Lines.Add(line);
            this.node.AddChild(
                new TreeNodeContentBuilder(this.node.Content).ResetLines()
                                                             .AddLine(line)
                                                             .SetNodeLabel("Line")
                                                             .SetHorizontalParagraph(horizontalLocation)
                                                             .Build());
        }
    }
}
