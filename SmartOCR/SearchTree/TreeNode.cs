using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartOCR
{
    class TreeNode
    {
        public List<TreeNode> Children { get; set; }
        public long HorizontalParagraph { get; set; }
        public List<long> Lines { get; set; }
        public string Name { get; set; }
        public string NodeLabel { get; set; }
        public TreeNode Parent { get; set; }
        public string RE_Pattern { get; set; }
        public bool Status { get; set; }
        public string Value { get; set; }
        public string ValueType { get; set; }

        public TreeNode() { }
        private TreeNode CreateNode(string new_name = "", string pattern = "", TreeNode parent_node = null, string value_type = "", long found_line = 0, string new_value = "", string node_label = "", long horizontal_paragraph = 0)
        {
            return new TreeNode()
            {
                Children = new List<TreeNode>(),
                Lines = new List<long>() { found_line },
                Name = new_name,
                NodeLabel = node_label,
                HorizontalParagraph = horizontal_paragraph,
                RE_Pattern = pattern,
                Parent = parent_node,
                Value = new_value,
                ValueType = value_type
            };
        }

        public TreeNode CreateRoot()
        {
            return CreateNode("root");
        }

        public TreeNode AddChild(string new_name = "", string pattern = "", long found_line = 0, string value_type = "", string new_value = "", string node_label = "", long horizontal_paragraph = 0)
        {
            TreeNode new_node = CreateNode(new_name, pattern, this, value_type, found_line, new_value, node_label, horizontal_paragraph);
            if (string.IsNullOrEmpty(new_node.Name))
                new_node.Name = Name;
            if (string.IsNullOrEmpty(new_node.ValueType))
                new_node.ValueType = ValueType;
            Children.Add(new_node);
            return new_node;
        }
        public TreeNode AddChild(TreeNode node)
        {
            if (node.Parent == null)
                node.Parent = this;
            Children.Add(node);
            return node;
        }
        public TreeNode AddSibling(long line_number = 0)
        {
            TreeNode new_node = CreateNode(Name, RE_Pattern, Parent, ValueType, line_number, Value, NodeLabel, HorizontalParagraph);
            if (!Parent.Lines.Contains(line_number))
                Parent.Lines.Add(line_number);
            return new_node;
        }
    }
}
