namespace SmartOCR.Parse
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using SmartOCR.Config;
    using SmartOCR.Search;
    using SmartOCR.Search.NodeProcessors;
    using SmartOCR.Utilities;
    using SmartOCR.Word;
    using SmartOCR.Word.Read;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

    /// <summary>
    /// Used to find data, specified by configuration fields, in Word document.
    /// </summary>
    public class WordParser
    {
        private ConfigData configData;
        private GridCollection gridCollection;
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
            Utilities.Debug("Parsing document.", 1);
            this.treeStructure.PopulateTree();
            this.ProcessDocument();
            return this.treeStructure.GetValuesFromTree();
        }

        private static bool DoesTreeNodeHasGridCoordinates(TreeNode fieldNode) => !Equals(fieldNode.Content.GridCoordinates, new Tuple<int, int>(-1, -1));

        private static bool ProcessTableNodeWithinGridSegment(TreeNode fieldNode, GridStructure gridStructure)
        {
            var nodeCopy = fieldNode.DeepCopy();
            new TableNodeProcessor(gridStructure[fieldNode.Content.GridCoordinates].Item2, fieldNode).Process();
            if (!fieldNode.Content.Status)
            {
                fieldNode = nodeCopy;
            }

            return fieldNode.Content.Status;
        }

        private void InitializeFields(WordReader reader, ConfigData data)
        {
            this.InitializeFieldsFromWordReader(reader);
            this.treeStructure = new SearchTree(data);
            this.configData = data;
        }

        private void InitializeFieldsFromWordReader(WordReader reader)
        {
            this.gridCollection = reader.GridCollection;
            this.lineMapping = reader.Mapping;
            this.tables = reader.TableCollection;
        }

        private void ProcessDocument()
        {
            foreach (var fieldNode in this.treeStructure.Children)
            {
                this.InitializeNodeAndGetDataByValueType(fieldNode);
            }
        }

        private void InitializeNodeAndGetDataByValueType(TreeNode fieldNode)
        {
            if (fieldNode.Content.ValueType.Contains("Table"))
            {
                Utilities.Debug($"Performing search for table node '{fieldNode.Content.Name}' data.", 2);
                this.InitializeNodeAndGetData(fieldNode, ProcessTableNodeWithinGridSegment, this.ProcessTableNodeWithinAllGrid);
            }
            else
            {
                this.InitializeNodeAndGetData(fieldNode, this.ProcessFieldNodeWithinGridSegment, this.ProcessFieldNodeWithinAllGrid);
            }
        }

        private void InitializeNodeAndGetData(TreeNode fieldNode, Func<TreeNode, GridStructure, bool> segmentSearchFunc, Action<TreeNode> postSegmentSearchAction)
        {
            if (DoesTreeNodeHasGridCoordinates(fieldNode))
            {
                if (this.gridCollection.Any(pageGridPair => segmentSearchFunc(fieldNode, pageGridPair.Value)))
                {
                    return;
                }
            }

            postSegmentSearchAction(fieldNode);
        }

        private void ProcessTableNodeWithinAllGrid(TreeNode fieldNode)
        {
            new TableNodeProcessor(this.tables, fieldNode).Process();
        }

        private void ProcessFieldNodeWithinAllGrid(TreeNode fieldNode)
        {
            new FieldNodeProcessor(this.configData[fieldNode.Content.Name], this.lineMapping).ProcessFieldNode(fieldNode);
        }

        private bool ProcessFieldNodeWithinGridSegment(TreeNode fieldNode, GridStructure gridStructure)
        {
            var nodeCopy = fieldNode.DeepCopy();
            new FieldNodeProcessor(
                this.configData[fieldNode.Content.Name],
                gridStructure[fieldNode.Content.GridCoordinates].Item1).ProcessFieldNode(fieldNode);
            if (!fieldNode.Content.Status)
            {
                fieldNode = nodeCopy;
            }

            return fieldNode.Content.Status;
        }
    }
}