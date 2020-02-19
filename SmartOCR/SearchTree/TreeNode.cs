using System;
using System.Collections.Generic;

namespace SmartOCR
{
    public class TreeNode // TODO: add summary.
    {
        public ITreeNodeContent Content { get; set; }
        public TreeNode Parent { get; set; }
        public List<TreeNode> Children { get; }

        public TreeNode()
        {
            Children = new List<TreeNode>();
            Content = new TreeNodeContent();
        }

        public TreeNode(TreeNode parentNode) : this()
        {
            Parent = parentNode;
        }

        public TreeNode(ITreeNodeContent content) : this()
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

        public static TreeNode CreateRoot()
        {
            return CreateNode("root");
        }

        public TreeNode AddChild(string newName = "", string pattern = "", long foundLine = 0, string valueType = "", string newValue = "", string nodeLabel = "", decimal horizontalParagraph = 0)
        {
            TreeNodeContent content = new TreeNodeContent()
            {
                Name = newName,
                RegExPattern = pattern,
                ValueType = valueType,
                CheckValue = newValue,
                NodeLabel = nodeLabel,
                HorizontalParagraph = horizontalParagraph
            };
            content.Lines.Add(foundLine);
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

        public TreeNode AddChild(TreeNode node, long newLine)
        {
            if (node == null)
            {
                throw new ArgumentNullException(nameof(node));
            }
            if (node.Content is TreeNodeContent)
            {
                TreeNodeContent node_content = (TreeNodeContent)node.Content;
                return AddChild(
                    node_content.Name, node_content.RegExPattern, newLine,
                    node_content.ValueType, node_content.CheckValue, node_content.NodeLabel,
                    node_content.HorizontalParagraph);
            }
            else
            {
                TableTreeNodeContent node_content = (TableTreeNodeContent)node.Content;
                return AddChild(
                    node_content.Name, node_content.RegExPattern, newLine,
                    node_content.ValueType, node_content.CheckValue, node_content.NodeLabel);
            }
        }

        public TreeNode AddChild(TreeNode node)
        {
            if (node == null)
            {
                throw new ArgumentNullException(nameof(node));
            }
            if (node.Parent == null)
            {
                node.Parent = this;
            }
            Children.Add(node);
            return node;
        }

        public TreeNode AddSibling(long lineNumber = 0)
        {
            TreeNode new_node = new TreeNode(this.Content);
            new_node.Content.Lines.Add(lineNumber);
            new_node.Parent = this.Parent;
            var parent_lines = Parent.Content.Lines;
            if (!parent_lines.Contains(lineNumber))
            {
                parent_lines.Add(lineNumber);
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
                RegExPattern = pattern,
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