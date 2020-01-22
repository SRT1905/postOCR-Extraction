using System.Collections.Generic;

namespace SmartOCR
{
    internal class TreeNode // TODO: add summary.
    {
        public TreeNodeContent Content;
        public TreeNode Parent;
        public List<TreeNode> Children;

        public TreeNode()
        {
            Children = new List<TreeNode>();
            Content = new TreeNodeContent();
        }

        public TreeNode(TreeNode parent_node) : this()
        {
            Parent = parent_node;
        }

        public TreeNode(TreeNodeContent content) : this()
        {
            this.Content = content;
        }

        private static TreeNode CreateNode(string new_name = "", string pattern = "", string value_type = "", long found_line = 0, string new_value = "", string node_label = "", long horizontal_paragraph = 0)
        {
            var new_node = new TreeNode
            {
                Content = PopulateNodeContent(new_name, pattern, value_type, found_line, new_value, node_label, horizontal_paragraph)
            };
            return new_node;
        }

        public TreeNode CreateRoot()
        {
            return CreateNode("root");
        }

        public TreeNode AddChild(string new_name = "", string pattern = "", long found_line = 0, string value_type = "", string new_value = "", string node_label = "", decimal horizontal_paragraph = 0)
        {
            TreeNodeContent content = new TreeNodeContent()
            {
                Name = new_name,
                RE_Pattern = pattern,
                ValueType = value_type,
                CheckValue = new_value,
                NodeLabel = node_label,
                HorizontalParagraph = horizontal_paragraph
            };
            content.Lines.Add(found_line);
            TreeNode new_node = new TreeNode(content)
            {
                Parent = this
            };
            if (string.IsNullOrEmpty(content.Name))
            {
                new_node.Content.Name = this.Content.Name;
            }

            if (string.IsNullOrEmpty(new_node.Content.ValueType))
            {
                new_node.Content.ValueType = this.Content.ValueType;
            }
            Children.Add(new_node);
            return new_node;
        }

        public TreeNode AddChild(TreeNode node, long new_line)
        {
            TreeNodeContent node_content = node.Content;
            return AddChild(
                node_content.Name, node_content.RE_Pattern, new_line,
                node_content.ValueType, node_content.CheckValue, node_content.NodeLabel,
                node_content.HorizontalParagraph);
        }

        public TreeNode AddChild(TreeNode node)
        {
            if (node.Parent == null)
            {
                node.Parent = this;
            }
            Children.Add(node);
            return node;
        }

        public TreeNode AddSibling(long line_number = 0)
        {
            TreeNode new_node = new TreeNode(this.Content);
            new_node.Content.Lines.Add(line_number);
            new_node.Parent = this.Parent;
            var parent_lines = Parent.Content.Lines;
            if (!parent_lines.Contains(line_number))
            {
                parent_lines.Add(line_number);
            }

            return new_node;
        }

        private static TreeNodeContent PopulateNodeContent(string new_name = "", string pattern = "", string value_type = "",
                                                    long found_line = 0, string new_value = "", string node_label = "",
                                                    long horizontal_paragraph = 0)
        {
            var content = new TreeNodeContent()
            {
                Name = new_name,
                RE_Pattern = pattern,
                CheckValue = new_value,
                ValueType = value_type,
                HorizontalParagraph = horizontal_paragraph,
                NodeLabel = node_label
            };
            content.Lines.Add(found_line);
            return content;
        }

        public override string ToString()
        {
            if (this.Content != null)
            {
                return $"{this.Content.Name}: {this.Content.NodeLabel} node";
            }
            return base.ToString();
        }
    }
}