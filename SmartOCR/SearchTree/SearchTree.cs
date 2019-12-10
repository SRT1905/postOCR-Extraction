using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    internal class SearchTree
    {
        private ConfigData config_data { get; }
        private TreeNode tree_structure;

        public SearchTree(ConfigData config_data)
        {
            this.config_data = config_data;
        }

        public List<TreeNode> Children
        {
            get
            {
                return tree_structure.Children;
            }
        }

        public void PopulateTree()
        {
            var root = new TreeNode().CreateRoot();
            foreach (ConfigField field in config_data.Fields)
            {
                TreeNode field_node = AddFieldNode(root, field);
                AddSearchValues(field, field_node);
            }
            tree_structure = root;
        }

        private TreeNode AddFieldNode(TreeNode root_node, ConfigField field_data)
        {
            var paragraph_collection = new List<long>() { 0 };

            TreeNodeContent content = new TreeNodeContent()
            {
                Name = field_data.Name,
                RE_Pattern = field_data.RE_Pattern,
                NodeLabel = "Field",
                ValueType = field_data.ValueType,
                CheckValue = field_data.ExpectedName,
            };
            content.Lines.Add(paragraph_collection[0]);

            TreeNode node = new TreeNode(content);

            for (int i = 0; i < paragraph_collection.Count; i++)
            {
                var child_content = new TreeNodeContent(content)
                {
                    NodeLabel = "Line"
                };
                child_content.Lines.Add(paragraph_collection[i]);
                var child_node = new TreeNode(child_content);
                node.AddChild(child_node);
            }
            root_node.AddChild(node);
            return node;
        }

        public void AddSearchValues(ConfigField field_data, TreeNode node, int initial_value_index = 0)
        {
            var node_content = node.Content;
            if (node_content.NodeLabel == "Terminal")
            {
                return;
            }

            List<ConfigExpression> values_collection = field_data.Expressions;
            if (node.Children.Count == 0 && values_collection.Count < initial_value_index + 1)
            {
                AddSearchValuesToChildlessNode(node, initial_value_index - 1, values_collection);
            }

            if (values_collection.Count < initial_value_index + 1)
            {
                return;
            }

            string field_name = node_content.Name;
            if (node_content.NodeLabel == "Line" || node_content.NodeLabel == "Field")
            {
                for (int i = 0; i < node.Children.Count; i++)
                {
                    TreeNode child = node.Children[i];
                    AddSearchValuesToSingleNode(field_name, child, values_collection, initial_value_index);
                }
            }
            else
            {
                AddSearchValuesToSingleNode(field_name, node, values_collection, initial_value_index);
            }
        }

        private void AddSearchValuesToChildlessNode(TreeNode node, int initial_value_index, List<ConfigExpression> values_collection)
        {
            ConfigExpression single_value_definition = values_collection[initial_value_index];
            var content = new TreeNodeContent()
            {
                Name = node.Content.Name,
                NodeLabel = $"Search {initial_value_index}",
                RE_Pattern = single_value_definition.RE_Pattern,
                HorizontalParagraph = node.Content.HorizontalParagraph,
                HorizontalStatus = single_value_definition.HorizontalStatus,
                ValueType = node.Content.ValueType
            };
            if (initial_value_index + 1 == values_collection.Count)
            {
                content.NodeLabel = "Terminal";
            }
            content.Lines.Add(node.Content.Lines[0] + single_value_definition.LineOffset);
            TreeNode new_node = new TreeNode(content);
            node.AddChild(new_node);
        }

        private void AddSearchValuesToSingleNode(string field_name, TreeNode node, List<ConfigExpression> values_collection, int initial_value_index)
        {
            TreeNode single_paragraph_node = node;
            for (int value_index = initial_value_index; value_index < values_collection.Count; value_index++)
            {
                ConfigExpression single_value_definition = values_collection[value_index];
                long offset_line = single_paragraph_node.Content.Lines[0];

                var content = new TreeNodeContent
                {
                    Name = field_name,
                    NodeLabel = $"Search {value_index}",
                    RE_Pattern = single_value_definition.RE_Pattern,
                    HorizontalParagraph = single_paragraph_node.Content.HorizontalParagraph,
                    ValueType = single_paragraph_node.Content.ValueType,
                    HorizontalStatus = single_value_definition.HorizontalStatus,
                };

                if (value_index + 1 == values_collection.Count)
                {
                    content.NodeLabel = "Terminal";
                }

                if (offset_line == 0)
                {
                    content.Lines.Add(0);
                }
                else
                {
                    content.Lines.Add(single_value_definition.LineOffset + offset_line);
                }

                var new_node = new TreeNode(content);
                single_paragraph_node = single_paragraph_node.AddChild(new_node);
            }
        }

        public Dictionary<string, string> GetValuesFromTree()
        {
            var final_values = new Dictionary<string, string>();
            foreach (ConfigField field in config_data.Fields)
            {
                List<string> children_collection = GetChildrenByFieldName(field.Name);
                var result = new HashSet<string>();
                if (children_collection.Count != 0)
                {
                    result.UnionWith(children_collection);
                }
                else
                {
                    var pre_terminal_collection = GetDataFromPreTerminalNodes(field.Name);
                    result.UnionWith(pre_terminal_collection);
                }
                final_values.Add(field.Name, string.Join("|", result));
            }
            return final_values;
        }

        private HashSet<string> GetDataFromPreTerminalNodes(string field_name)
        {
            var found_data = new Dictionary<bool, HashSet<string>>()
            {
                { true, new HashSet<string>() },
                { false, new HashSet<string>() }
            };
            foreach (TreeNode node in tree_structure.Children)
            {
                if (node.Content.Name == field_name)
                {
                    GetDataFromNode(node, found_data);
                    if (found_data[true].Count != 0)
                    {
                        return found_data[true];
                    }
                    return found_data[false];
                    
                }
            }
            return new HashSet<string>();
        }

        private void GetDataFromNode(TreeNode node, Dictionary<bool, HashSet<string>> found_data)
        {
            if (node.Children.Count == 0)
            {
                return;
            }

            foreach (TreeNode child in node.Children)
            {
                GetDataFromNode(child, found_data);
            }
            if (!string.IsNullOrEmpty(node.Content.FoundValue))
            {
                if (node.Content.Status)
                {
                    found_data[true].Add(node.Content.FoundValue);
                }
                else
                {
                    if (node.Parent.Content.NodeLabel != "Field")
                    {
                        found_data[false].Add(node.Content.FoundValue);
                    }
                }

            }

        }

        private List<string> GetChildrenByFieldName(string field_name)
        {
            var children_collection = new List<string>();
            foreach (TreeNode field_node in tree_structure.Children)
            {
                if (field_node.Content.Name == field_name)
                {
                    GetNodeChildren(field_node, children_collection);
                    break;
                }
            }

            return children_collection;
        }

        private void GetNodeChildren(TreeNode node, List<string> children_collection)
        {
            if (node.Children.Count == 0)
            {
                if (node.Content.Status)
                {
                    children_collection.Add(node.Content.FoundValue);
                }
                return;
            }

            foreach (TreeNode child in node.Children)
            {
                GetNodeChildren(child, children_collection);
            }
        }
    }
}