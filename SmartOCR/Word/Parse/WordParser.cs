namespace SmartOCR
{
    using System;
    using System.Collections.Generic;

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
            Utilities.Debug($"Parsing document.", 1);
            this.treeStructure.PopulateTree();
            this.ProcessDocument();
            return this.treeStructure.GetValuesFromTree();
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
                this.InitializeTableNodeAndGetData(fieldNode);
            }
            else
            {
                this.InitializeFieldNodeAndGetData(fieldNode);
            }
        }

        private void InitializeFieldNodeAndGetData(TreeNode fieldNode)
        {
            if (!Equals(fieldNode.Content.GridCoordinates, new Tuple<int, int>(-1, -1)))
            {
                foreach (var pageGridPair in this.gridCollection)
                {
                    if (this.ProcessFieldNodeWithinGridSegment(fieldNode, pageGridPair.Value))
                    {
                        return;
                    }
                }
            }

            this.ProcessFieldNodeWithinAllGrid(fieldNode);
        }

        private void ProcessFieldNodeWithinAllGrid(TreeNode fieldNode)
        {
            new FieldNodeProcessor(this.configData[fieldNode.Content.Name], this.lineMapping).ProcessFieldNode(fieldNode);
        }

        private bool ProcessFieldNodeWithinGridSegment(TreeNode fieldNode, GridStructure gridStructure)
        {
            new FieldNodeProcessor(this.configData[fieldNode.Content.Name], gridStructure[fieldNode.Content.GridCoordinates]).ProcessFieldNode(fieldNode);
            if (!fieldNode.Content.Status)
            {
                fieldNode.Reset();
            }

            return fieldNode.Content.Status;
        }

        private void InitializeTableNodeAndGetData(TreeNode fieldNode)
        {
            Utilities.Debug($"Performing search for table node '{fieldNode.Content.Name}' data.", 2);
            new TableNodeProcessor(this.tables, fieldNode).Process();
        }
    }
}