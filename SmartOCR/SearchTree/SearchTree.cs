using System.Collections;
using System.Collections.Generic;

namespace SmartOCR
{
    internal class SearchTree
    {
        private Dictionary<string, object> config_data { get; }
        private TreeNode tree_structure;

        public SearchTree(Dictionary<string, object> config_data)
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
            foreach (string key in config_data.Keys)
            {
                TreeNode field_node = AddFieldNode(root, key, config_data[key]);
                AddSearchValues(config_data[key], field_node);
            }
            tree_structure = root;
        }

        private TreeNode AddFieldNode(TreeNode root_node, string field_name, object field_data)
        {
            var field_info = (Dictionary<string, object>)field_data;
            var field_expression = (Dictionary<string, object>)field_info["field_expression"];
            var paragraph_collection = (List<long>)field_info["paragraphs"];
            string value_type = (string)field_info["type"];

            if (paragraph_collection.Count == 0)
            {
                TreeNodeContent new_content = new TreeNodeContent()
                {
                    Name = field_name,
                    RE_Pattern = (string)field_expression["regexp"],
                    ValueType = value_type,
                    CheckValue = (string)field_expression["value"],
                    NodeLabel = "Field"
                };
                new_content.Lines.Add(0);
                var tree_node = new TreeNode(new_content);
                return root_node.AddChild(tree_node);
            }

            TreeNodeContent content = new TreeNodeContent()
            {
                Name = field_name,
                RE_Pattern = (string)field_expression["regexp"],
                NodeLabel = "Field",
                ValueType = value_type,
                CheckValue = (string)field_expression["value"]
            };
            content.Lines.Add(paragraph_collection[0]);

            TreeNode node = new TreeNode(content);

            for (int i = 0; i < paragraph_collection.Count; i++)
            {
                var child_content = new TreeNodeContent()
                {
                    Name = field_name,
                    RE_Pattern = (string)field_expression["regexp"],
                    NodeLabel = "Line",
                    ValueType = value_type,
                    CheckValue = (string)field_expression["value"]
                };
                child_content.Lines.Add(paragraph_collection[i]);
                var child_node = new TreeNode(child_content);
                node.AddChild(child_node);
            }
            root_node.AddChild(node);
            return node;
        }

        public void AddSearchValues(dynamic field_data, TreeNode node, int initial_value_index = 0)
        {
            var node_content = node.Content;
            if (node_content.NodeLabel == "Terminal")
            {
                return;
            }

            ArrayList values_collection = field_data["values"];
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

        private void AddSearchValuesToChildlessNode(TreeNode node, int initial_value_index, ArrayList values_collection)
        {
            Dictionary<string, object> single_value_definition = (Dictionary<string, object>)values_collection[initial_value_index];
            var content = new TreeNodeContent()
            {
                Name = node.Content.Name,
                NodeLabel = $"Search {initial_value_index}",
                RE_Pattern = (string)single_value_definition["regexp"],
                HorizontalParagraph = node.Content.HorizontalParagraph == 0
                        ? node.Content.HorizontalParagraph
                        : (long)single_value_definition["horizontal_offset"] + node.Content.HorizontalParagraph,
                ValueType = node.Content.ValueType
            };
            if (initial_value_index + 1 == values_collection.Count)
            {
                content.NodeLabel = "Terminal";
            }
            content.Lines.Add(node.Content.Lines[0] + (long)single_value_definition["offset"]);
            TreeNode new_node = new TreeNode(content);
            node.AddChild(new_node);
        }

        private void AddSearchValuesToSingleNode(string field_name, TreeNode node, ArrayList values_collection, int initial_value_index)
        {
            TreeNode single_paragraph_node = node;
            for (int value_index = initial_value_index; value_index < values_collection.Count; value_index++)
            {
                Dictionary<string, dynamic> single_value_definition = (Dictionary<string, dynamic>)values_collection[value_index];
                long offset_line = single_paragraph_node.Content.Lines[0];
                long node_paragraph = single_paragraph_node.Content.HorizontalParagraph;

                var content = new TreeNodeContent
                {
                    Name = field_name,
                    NodeLabel = $"Search {value_index}",
                    RE_Pattern = (string)single_value_definition["regexp"],
                    HorizontalParagraph = node_paragraph == 0
                        ? node_paragraph
                        : (long)single_value_definition["horizontal_offset"] + node_paragraph,
                    ValueType = single_paragraph_node.Content.ValueType,
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
                    content.Lines.Add(single_value_definition["offset"] + offset_line);
                }

                var new_node = new TreeNode(content);
                single_paragraph_node = single_paragraph_node.AddChild(new_node);
            }
        }

        public Dictionary<string, string> GetValuesFromTree()
        {
            var final_values = new Dictionary<string, string>();
            foreach (string field_name in config_data.Keys)
            {
                List<string> children_collection = GetChildrenByFieldName(field_name);
                var result = new HashSet<string>();
                if (children_collection.Count != 0)
                {
                    result.UnionWith(children_collection);
                }
                else
                {
                    var pre_terminal_collection = GetDataFromPreTerminalNodes(field_name);
                    result.UnionWith(pre_terminal_collection);
                }
                final_values.Add(field_name, string.Join("|", result));
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