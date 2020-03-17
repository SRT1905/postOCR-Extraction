namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Used to find data, specified by configuration fields, in Word document.
    /// </summary>
    public class WordParser
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

        private void InitializeFields(WordReader reader, ConfigData configData)
        {
            this.InitializeFieldsFromWordReader(reader);
            this.treeStructure = new SearchTree(configData);
            this.configData = configData;
        }

        private void InitializeFieldsFromWordReader(WordReader reader)
        {
            this.lineMapping = reader.Mapping;
            this.tables = reader.TableCollection;
        }

        private void ProcessDocument()
        {
            for (int fieldIndex = 0; fieldIndex < this.treeStructure.Children.Count; fieldIndex++)
            {
                this.InitializeNodeAndGetDataByValueType(this.treeStructure.Children[fieldIndex]);
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
                this.InitializeLineNodeAndGetData(fieldNode);
            }
        }

        private void InitializeLineNodeAndGetData(TreeNode fieldNode)
        {
            fieldNode = this.InitializeFieldNode(fieldNode);

            if (fieldNode.Content.Lines[0] != 0)
            {
                Utilities.Debug($"Performing search for field node '{fieldNode.Content.Name}' data.", 2);
                this.ProcessFieldNodeChildren(fieldNode);
            }
        }

        private void ProcessFieldNodeChildren(TreeNode fieldNode)
        {
            foreach (TreeNode lineNode in fieldNode.Children)
            {
                this.ProcessSingleLineNode(fieldNode, lineNode);
            }
        }

        private void ProcessSingleLineNode(TreeNode fieldNode, TreeNode lineNode)
        {
            new LineNodeProcessor(this.configData[fieldNode.Content.Name], this.lineMapping)
                                    .ProcessLineNode(lineNode);
        }

        private void InitializeTableNodeAndGetData(TreeNode fieldNode)
        {
            Utilities.Debug($"Performing search for table node '{fieldNode.Content.Name}' data.", 2);
            new TableNodeProcessor(this.tables, fieldNode).Process();
        }

        private TreeNode InitializeFieldNode(TreeNode fieldNode)
        {
            if (fieldNode.Content.Lines[0] == 0)
            {
                Utilities.Debug($"Initializing field node '{fieldNode.Content.Name}' data.", 2);
                fieldNode = new UndefinedNodeProcessor(fieldNode, this.lineMapping, this.configData).GetProcessedNode();
            }

            return fieldNode;
        }
    }
}