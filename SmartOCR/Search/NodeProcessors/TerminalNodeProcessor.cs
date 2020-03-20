namespace SmartOCR.Search.NodeProcessors
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using SmartOCR.Config;
    using SmartOCR.Word;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

    /// <summary>
    /// Used to process values, specified by configuration fields, in Word paragraphs, distributed by lines.
    /// </summary>
    public class TerminalNodeProcessor
    {
        private readonly ConfigField configField;
        private readonly LineMapping lineMapping;

        /// <summary>
        /// Initializes a new instance of the <see cref="TerminalNodeProcessor"/> class.
        /// </summary>
        /// <param name="configField">A source description of search field.</param>
        /// <param name="lineMapping">A mapping between document line and paragraphs on it.</param>
        public TerminalNodeProcessor(ConfigField configField, LineMapping lineMapping)
        {
            this.configField = configField;
            this.lineMapping = lineMapping;
        }

        /// <summary>
        /// Performs processing of value, described by <paramref name="terminalNode"/>.
        /// </summary>
        /// <param name="terminalNode">An instance of <see cref="TreeNode"/> class, in respect to which search is made.</param>
        /// <param name="searchLevel">Level of search, equals 0 by default.</param>
        public void Process(TreeNode terminalNode, int searchLevel)
        {
            Utilities.Debug($"Processing terminal node '{terminalNode.Content.Name}'.", 3);
            foreach (int lineNumber in terminalNode.Content.Lines)
            {
                if (this.lineMapping.ContainsKey(lineNumber))
                {
                    if (this.ProcessDataInLine(terminalNode, lineNumber))
                    {
                        return;
                    }

                    this.DoOffsetSearch(terminalNode, searchLevel, lineNumber);
                }
                else
                {
                    this.DoOffsetSearch(terminalNode, searchLevel, lineNumber);
                    return;
                }
            }
        }

        private static void PropagateStatusUpTree(bool status, TreeNode node)
        {
            var tempNode = node;
            while (tempNode.Content.Name != "root")
            {
                tempNode.Content.Status = status;
                tempNode = tempNode.Parent;
            }
        }

        private static int GetParagraphByLocation(List<ParagraphContainer> paragraphCollection, decimal position, bool returnNextLargest)
        {
            int location = paragraphCollection
                .Select(item => item.HorizontalLocation)
                .ToList()
                .BinarySearch(position);
            return GetValidatedParagraphLocation(
                paragraphCollection,
                ValidateNegativeParagraphLocation(location),
                returnNextLargest);
        }

        private static int ValidateNegativeParagraphLocation(int location) => location < 0 ? ~location : location;

        private static int GetValidatedParagraphLocation(
            List<ParagraphContainer> paragraphCollection,
            int location,
            bool returnNextLargest) => !returnNextLargest || location == paragraphCollection.Count ? --location : location;

        private static void DoProceduresOnSuccessfulTerminalMatch(TreeNode node, string result)
        {
            Utilities.Debug($"Found successful match for terminal node '{node.Content.Name}': {result}.", 4);
            node.Content.FoundValue = result;
            PropagateStatusUpTree(true, node);
        }

        private static void ValidateMatchProcessorResult(TreeNode node, MatchProcessor matchProcessor)
        {
            if (!string.IsNullOrEmpty(matchProcessor.Result))
            {
                DoProceduresOnSuccessfulTerminalMatch(node, matchProcessor.Result);
            }
        }

        private static Tuple<int, int> DefineSearchIndexes(TreeNodeContent nodeContent, List<ParagraphContainer> paragraphCollection)
        {
            int startIndex = nodeContent.SecondSearchParameter == 1
                ? GetParagraphByLocation(paragraphCollection, nodeContent.HorizontalParagraph, true)
                : 0;
            int finishIndex = nodeContent.SecondSearchParameter == -1
                ? GetParagraphByLocation(paragraphCollection, nodeContent.HorizontalParagraph, false)
                : paragraphCollection.Count - 1;

            return Tuple.Create(startIndex, finishIndex);
        }

        private bool ProcessDataInLine(TreeNode terminalNode, int lineNumber)
        {
            List<ParagraphContainer> paragraphCollection = this.lineMapping[lineNumber];
            var indexTuple = DefineSearchIndexes(terminalNode.Content, paragraphCollection);

            for (int paragraphIndex = indexTuple.Item1; paragraphIndex <= indexTuple.Item2; paragraphIndex++)
            {
                if (this.ProcessValueInSingleParagraph(paragraphCollection[paragraphIndex].Text, terminalNode))
                {
                    return true;
                }
            }

            return false;
        }

        private void DoOffsetSearch(TreeNode terminalNode, int searchLevel, int lineNumber)
        {
            new OffsetNodeProcessor(this.configField, this.lineMapping)
                .OffsetSearch(terminalNode, lineNumber, searchLevel, addToParentStatus: true);
        }

        private bool ProcessValueInSingleParagraph(string paragraphText, TreeNode node)
        {
            var matchProcessor = new MatchProcessor(
                paragraphText,
                Utilities.CreateRegexpObject(node.Content.TextExpression),
                node.Content.ValueType);
            ValidateMatchProcessorResult(node, matchProcessor);

            return node.Content.Status;
        }
    }
}
