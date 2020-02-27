namespace SmartOCR
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Used to contain information about search status of specific config field.
    /// </summary>
    public class TreeNode
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TreeNode"/> class.
        /// </summary>
        public TreeNode()
        {
            this.Children = new List<TreeNode>();
            this.Content = new TreeNodeContent();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TreeNode"/> class.
        /// Instance contents are initialized by provided instance of <see cref="TreeNodeContent"/> class.
        /// </summary>
        /// <param name="content">An instance of <see cref="TreeNodeContent"/> class with information about existing<see cref="TreeNode"/> instance.</param>
        public TreeNode(TreeNodeContent content)
            : this()
        {
            this.Content = content;
        }

        /// <summary>
        /// Gets or sets information about search status of specific config field.
        /// </summary>
        public TreeNodeContent Content { get; set; }

        /// <summary>
        /// Gets collection of <see cref="TreeNode"/> instance with deeper search status.
        /// </summary>
        public List<TreeNode> Children { get; }

        /// <summary>
        /// Gets or sets a higher level instance of <see cref="TreeNode"/> class.
        /// </summary>
        public TreeNode Parent { get; set; }

        /// <summary>
        /// Creates an empty tree structure with a single root node.
        /// </summary>
        /// <returns>An instance of <see cref="TreeNode"/> named "root".</returns>
        public static TreeNode CreateRoot()
        {
            var builder = new TreeNodeContentBuilder();
            builder.SetName("root");
            builder.TryAddLine(0);
            builder.SetHorizontalParagraph(0);
            return new TreeNode(builder.Build());
        }

        /// <summary>
        /// Adds provided node to collection of lower level nodes.
        /// </summary>
        /// <param name="node">An instance of <see cref="TreeNode"/> class that is added to current node.</param>
        /// <returns>Provided instance of <see cref="TreeNode"/> that has current node as its parent.</returns>
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

        /// <summary>
        /// Creates a new instance of <see cref="TreeNode"/> class, initialized by <see cref="TreeNodeContent"/> instance, and adds new node to collection of lower level nodes.
        /// </summary>
        /// <param name="content">An instance of <see cref="TreeNodeContent"/> class that is used to create new node.</param>
        /// <returns>Created instance of <see cref="TreeNode"/> that has current node as its parent.</returns>
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

        /// <summary>
        /// Returns string representation of current <see cref="TreeNode"/> instance.
        /// </summary>
        /// <returns>String representation of <see cref="TreeNode"/>instance.</returns>
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