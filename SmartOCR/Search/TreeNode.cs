namespace SmartOCR
{
    using System;
    using System.Collections.Generic;

    public class TreeNode // TODO: add summary.
    {
        public TreeNode()
        {
            this.Children = new List<TreeNode>();
            this.Content = new TreeNodeContent();
        }

        public TreeNode(TreeNodeContent content)
            : this()
        {
            this.Content = content;
        }

        public TreeNodeContent Content { get; set; }

        public List<TreeNode> Children { get; }

        public TreeNode Parent { get; set; }

        public static TreeNode CreateRoot()
        {
            var builder = new TreeNodeContentBuilder();
            builder.SetName("root");
            builder.TryAddLine(0);
            builder.SetHorizontalParagraph(0);
            return new TreeNode(builder.Build());
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

            this.Children.Add(node);
            return node;
        }

        public TreeNode AddChild(TreeNodeContent content)
        {
            if (string.IsNullOrEmpty(content.Name))
            {
                content.Name = this.Content.Name;
            }

            if (string.IsNullOrEmpty(content.ValueType))
            {
                content.ValueType = this.Content.ValueType;
            }

            TreeNode node = new TreeNode(content)
            {
                Parent = this,
            };
            this.Children.Add(node);
            return node;
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