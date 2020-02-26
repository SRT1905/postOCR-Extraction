using System;
using System.Collections.Generic;

namespace SmartOCR
{
    public class TreeNode // TODO: add summary.
    {
        #region Properties
        public TreeNodeContent Content { get; set; }
        public List<TreeNode> Children { get; }
        public TreeNode Parent { get; set; }
        #endregion

        #region Constructors
        public TreeNode()
        {
            Children = new List<TreeNode>();
            Content = new TreeNodeContent();
        }
        public TreeNode(TreeNodeContent content) : this()
        {
            this.Content = content;
        }
        #endregion

        #region Public methods
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
        public TreeNode AddChild(TreeNodeContent content)
        {
            if (string.IsNullOrEmpty(content.Name))
            {
                content.Name = Content.Name;
            }

            if (string.IsNullOrEmpty(content.ValueType))
            {
                content.ValueType = Content.ValueType;
            }
            TreeNode node = new TreeNode(content)
            {
                Parent = this
            };
            Children.Add(node);
            return node;
        }
        public override string ToString()
        {
            if (Content != null)
            {
                return $"{Content.Name}: {Content.NodeLabel} node";
            }
            return base.ToString();
        }
        #endregion

        #region Public static methods
        public static TreeNode CreateRoot()
        {
            var builder = new TreeNodeContentBuilder();
            builder.SetName("root");
            builder.TryAddLine(0);
            builder.SetHorizontalParagraph(0);
            return new TreeNode(builder.Build());
        }
        #endregion
    }
}