using System.Collections;
using System.Collections.Generic;

namespace SmartOCR
{
    internal class SearchTree
    {
        private readonly Dictionary<string, object> config_data;
        private TreeNode tree_structure;

        public SearchTree(Dictionary<string, object> config_data)
        {
            this.config_data = config_data;
            PopulateTree();
        }

        public List<TreeNode> Children => tree_structure.Children;

        private void PopulateTree()
        {
            TreeNode root = new TreeNode().CreateRoot();
            foreach (string key in config_data.Keys)
            {
                TreeNode field_node = AddFieldNode(ref root, key, config_data[key]);
                AddSearchValues(config_data[key], ref field_node);
            }
            tree_structure = root;
        }

        private TreeNode AddFieldNode(ref TreeNode root_node, string field_name, object field_data)
        {
            Dictionary<string, object> field_info = (Dictionary<string, object>)field_data;

            Dictionary<string, object> field_expression = (Dictionary<string, object>)field_info["field_expression"];
            List<long> paragraph_collection = (List<long>)field_info["paragraphs"];
            string value_type = (string)field_info["type"];

            if (paragraph_collection.Count == 0)
                return root_node.AddChild(field_name, (string)field_expression["regexp"], 0, value_type, (string)field_expression["value"], "Field");

            TreeNode node = new TreeNode()
            {
                Children = new List<TreeNode>() { },
                Lines = new List<long>() { paragraph_collection[0] },
                Name = field_name,
                RE_Pattern = (string)field_expression["regexp"],
                NodeLabel = "Field",
                ValueType = value_type,
                Value = (string)field_expression["value"]
            };
            node.AddChild(field_name, (string)field_expression["regexp"], paragraph_collection[0], value_type, (string)field_expression["value"], "Line");
            for (int i = 1; i < paragraph_collection.Count; i++)
            {
                node.Lines.Add(paragraph_collection[i]);
                node.AddChild(field_name, (string)field_expression["regexp"], paragraph_collection[i], value_type, (string)field_expression["value"], "Line");
            }
            root_node.AddChild(node);
            return node;
        }

        public void AddSearchValues(dynamic field_data, ref TreeNode node, int initial_value_index = 0)
        {
            if (node.NodeLabel == "Terminal")
            {
                return;
            }

            ArrayList values_collection = field_data["values"];
            if (node.Children.Count == 0 && values_collection.Count < initial_value_index + 1)
            {
                AddSearchValuesToChildlessNode(ref node, initial_value_index, values_collection);
            }

            if (values_collection.Count < initial_value_index + 1)
            {
                return;
            }

            string field_name = node.Name;
            if (node.NodeLabel == "Line" || node.NodeLabel == "Field")
            {
                for (int i = 0; i < node.Children.Count; i++)
                {
                    TreeNode child = node.Children[i];
                    AddSearchValuesToSingleNode(field_name, ref child, values_collection, initial_value_index);
                }
            }
            else
            {
                AddSearchValuesToSingleNode(field_name, ref node, values_collection, initial_value_index);
            }
        }

        private void AddSearchValuesToChildlessNode(ref TreeNode node, int initial_value_index,
                                                    ArrayList values_collection)
        {
            TreeNode new_node = new TreeNode();
            Dictionary<string, object> single_value_definition = (Dictionary<string, object>)values_collection[initial_value_index];
            new_node.RE_Pattern = (string)single_value_definition["regexp"];
            new_node.Lines.Add(node.Lines[0] + (long)single_value_definition["offset"]);
            node.AddChild(new_node);
        }

        private void AddSearchValuesToSingleNode(string field_name, ref TreeNode node, ArrayList values_collection, int initial_value_index)
        {
            TreeNode single_paragraph_node = node;
            for (int value_index = initial_value_index; value_index < values_collection.Count; value_index++)
            {
                string node_label = $"Search {value_index}";
                if (value_index + 1 == values_collection.Count)
                    node_label = "Terminal";
                Dictionary<string, dynamic> single_value_definition = (Dictionary<string, dynamic>)values_collection[value_index];
                string regexp_pattern = (string)single_value_definition["regexp"];
                long offset_line = single_paragraph_node.Lines[0];
                long node_paragraph = single_paragraph_node.HorizontalParagraph;
                long value_offset;
                if (offset_line == 0)
                    value_offset = 0;
                else
                    value_offset = single_value_definition["offset"] + offset_line;

                long horizontal_value_offset;
                if (node_paragraph == 0)
                    horizontal_value_offset = 0;
                else
                    horizontal_value_offset = (long)single_value_definition["horizontal_offset"] + node_paragraph;

                single_paragraph_node = single_paragraph_node.AddChild(field_name, regexp_pattern, value_offset, node_label: node_label, horizontal_paragraph: horizontal_value_offset);
            }
        }

        public Dictionary<string, string> GetValuesFromTree()
        {
            Dictionary<string, string> final_values = new Dictionary<string, string>();
            foreach (string field_name in config_data.Keys)
            {
                List<string> children_collection = GetChildrenByFieldName(field_name);

                if (children_collection.Count != 0)
                {
                    HashSet<string> result = new HashSet<string>(children_collection);
                    final_values.Add(field_name, string.Join("|", result));
                }
            }
            return final_values;
        }

        private List<string> GetChildrenByFieldName(string field_name)
        {
            List<string> children_collection = new List<string>();
            foreach (TreeNode field_node in tree_structure.Children)
            {
                if (field_node.Name == field_name)
                {
                    GetNodeChildren(field_node, ref children_collection);
                    break;
                }
            }

            return children_collection;
        }

        private void GetNodeChildren(TreeNode node, ref List<string> children_collection)
        {
            if (node.Children.Count == 0)
            {
                if (node.Status)
                    children_collection.Add(node.Value);
                return;
            }

            foreach (TreeNode child in node.Children)
            {
                GetNodeChildren(child, ref children_collection);
            }
        }
    }
}