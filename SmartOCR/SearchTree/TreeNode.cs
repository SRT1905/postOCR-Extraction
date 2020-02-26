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
        public TreeNode(TreeNode parentNode) : this()
        {
            Parent = parentNode;
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
        public TreeNode AddChild(TreeNode node, long newLine)
        {
            if (node == null)
            {
                throw new ArgumentNullException(nameof(node));
            }
            TreeNodeContent nodeContent = node.Content;
            return AddChild(nodeContent.Name, nodeContent.RegExPattern, newLine,
                            nodeContent.ValueType, nodeContent.CheckValue, nodeContent.NodeLabel,
                            nodeContent.HorizontalParagraph);
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
            TreeNode newNode = new TreeNode(content)
            {
                Parent = this
            };
            if (string.IsNullOrEmpty(content.Name))
            {
                newNode.Content.Name = this.Content.Name;
            }

            if (string.IsNullOrEmpty(newNode.Content.ValueType))
            {
                newNode.Content.ValueType = this.Content.ValueType;
            }
            Children.Add(newNode);
            return newNode;
        }
        public TreeNode AddSibling(long lineNumber = 0)
        {
            TreeNode newNode = new TreeNode(this.Content);
            newNode.Content.Lines.Add(lineNumber);
            newNode.Parent = this.Parent;
            var parentLines = Parent.Content.Lines;
            if (!parentLines.Contains(lineNumber))
            {
                parentLines.Add(lineNumber);
            }

            return newNode;
        }
        public override string ToString()
        {
            if (this.Content != null)
            {
                return $"{this.Content.Name}: {this.Content.NodeLabel} node";
            }
            return base.ToString();
        }
        #endregion

        #region Public static methods
        public static TreeNode CreateRoot()
        {
            return CreateNode("root");
        }
        #endregion

        #region Private static methods
        private static TreeNode CreateNode(string newName = "", string pattern = "",
                                           string valueType = "", long foundLine = 0,
                                           string newValue = "", string nodeLabel = "",
                                           long horizontalParagraph = 0)
        {
            var newNode = new TreeNode
            {
                Content = PopulateNodeContent(newName, pattern, valueType, foundLine, newValue, nodeLabel, horizontalParagraph)
            };
            return newNode;
        }


        private static TreeNodeContent PopulateNodeContent(string newName = "", string pattern = "", string valueType = "",
                                                           long foundLine = 0, string newValue = "", string nodeLabel = "",
                                                           long horizontalParagraph = 0)
        {
            var content = new TreeNodeContent()
            {
                Name = newName,
                RegExPattern = pattern,
                CheckValue = newValue,
                ValueType = valueType,
                HorizontalParagraph = horizontalParagraph,
                NodeLabel = nodeLabel
            };
            content.Lines.Add(foundLine);
            return content;
        }
        #endregion
    }
}