using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    internal class WordParser
    {
        private const long similarity_search_threshold = 8;
        private SearchTree tree_structure { get; }
        private SortedDictionary<long, List<ParagraphContainer>> line_mapping { get; }
        private Dictionary<string, object> config_data { get; }

        public WordParser(SortedDictionary<long, List<ParagraphContainer>> document_content, Dictionary<string, object> config_data)
        {
            this.line_mapping = document_content;
            this.tree_structure = new SearchTree(config_data);
            this.config_data = config_data;
        }

        public Dictionary<string, string> ParseDocument()
        {
            tree_structure.PopulateTree();
            ProcessDocument();
            return tree_structure.GetValuesFromTree();
        }

        private void AddChildrenToFieldNode(TreeNode field_node, Dictionary<string, SimilarityDescription> collected_data, double max_similarity)
        {
            foreach (string key in collected_data.Keys)
            {
                if (max_similarity == collected_data[key].Ratio)
                {
                    AddSingleChildToFieldNode(field_node, key);
                }
            }
        }

        private void AddSingleChildToFieldNode(TreeNode field_node, string key)
        {
            var content = field_node.Content;
            long line = long.Parse(key.Split('|')[0]);
            if (!content.Lines.Contains(line))
            {
                content.Lines.Add(line);
                field_node.AddChild(found_line: line, pattern: content.RE_Pattern, new_value: content.CheckValue, node_label: "Line");
            }
        }

        private void GetDataFromUndefinedNode(TreeNode field_node)
        {
            Regex regex_object = Utilities.CreateRegexpObject(field_node.Content.RE_Pattern);
            Dictionary<string, SimilarityDescription> collected_data = new Dictionary<string, SimilarityDescription>();

            foreach (long line in line_mapping.Keys)
            {
                foreach (ParagraphContainer container in line_mapping[line])
                {
                    if (regex_object.IsMatch(container.Text))
                    {
                        List<SimilarityDescription> matched_data_collection = GetMatchesFromParagraph(container.Text, regex_object, field_node.Content.CheckValue);

                        for (int i = 0; i < matched_data_collection.Count; i++)
                        {
                            collected_data.Add($"{line}|{line_mapping[line].IndexOf(container)}|{i}", matched_data_collection[i]);
                        }
                    }
                }
            }

            if (collected_data.Count != 0)
            {
                UpdateFieldNode(collected_data, field_node);
            }
        }

        private List<SimilarityDescription> GetMatchesFromParagraph(string text_to_check, Regex re_object, string check_value)
        {
            MatchCollection matches = re_object.Matches(text_to_check);
            List<SimilarityDescription> found_values = new List<SimilarityDescription>();
            foreach (Match single_match in matches)
            {
                if (single_match.Groups.Count > 1)
                {
                    for (int group_i = 1; group_i < single_match.Groups.Count; group_i++)
                    {
                        Group group_item = single_match.Groups[group_i];
                        SimilarityDescription description = new SimilarityDescription(group_item.Value, check_value);
                        if (description.CheckStringSimilarity())
                        {
                            found_values.Add(description);
                        }
                    }
                }
                else
                {
                    SimilarityDescription description = new SimilarityDescription(single_match.Value, check_value);
                    if (description.CheckStringSimilarity())
                    {
                        found_values.Add(description);
                    }
                }
            }
            return found_values;
        }

        private List<long> GetOffsetLines(long line_number, TreeNodeContent content)
        {
            Regex regex_obj = Utilities.CreateRegexpObject(content.RE_Pattern);
            List<long> found_values_collection = new List<long>();

            for (int search_offset = 1; search_offset <= similarity_search_threshold; search_offset++)
            {
                List<long> offset_indexes = new List<long>() { line_number + search_offset, line_number - search_offset };
                foreach (long offset_index in offset_indexes)
                {
                    if (line_mapping.ContainsKey(offset_index))
                    {
                        var line_checker = new LineContentChecker(line_mapping[offset_index]);
                        if (line_checker.CheckLineContents(regex_obj, content.CheckValue))
                        {
                            found_values_collection.Add(offset_index);
                        }
                    }
                }
            }
            return found_values_collection;
        }

        private void OffsetSearch(long line_number, TreeNode line_node, long search_level, bool add_to_parent = false)
        {
            var line_node_content = line_node.Content;
            List<long> line_numbers = GetOffsetLines(line_number, line_node_content);

            foreach (long offset_index in line_numbers)
            {
                if (add_to_parent)
                {
                    var parent = line_node.Parent;
                    AddOffsetNode(parent, search_level, offset_index, add_to_parent);
                }
                else
                {
                    AddOffsetNode(line_node, search_level, offset_index, add_to_parent);
                }
            }
        }

        private void AddOffsetNode(TreeNode node, long search_level, long offset_index, bool add_to_parent)
        {
            var node_content = node.Content;
            if (!node_content.Lines.Contains(offset_index))
            {
                node_content.Lines.Add(offset_index);
                string node_label;
                string pattern;
                if (add_to_parent)
                {
                    var first_child_content = node.Children.First().Content;
                    node_label = first_child_content.NodeLabel;
                    pattern = first_child_content.RE_Pattern;
                }
                else
                {
                    node_label = node.Content.NodeLabel;
                    pattern = node.Content.RE_Pattern;
                }
                TreeNode child_node = node.AddChild(found_line: offset_index, pattern: pattern, node_label: node_label);
                tree_structure.AddSearchValues(config_data[node_content.Name], child_node, (int)search_level);
            }
        }

        private void ProcessDocument()
        {
            for (int field_index = 0; field_index < tree_structure.Children.Count; field_index++)
            {
                TreeNode field_node = tree_structure.Children[field_index];
                var node_content = field_node.Content;
                if (node_content.Lines[0] == 0)
                {
                    GetDataFromUndefinedNode(field_node);
                }
                if (node_content.Lines[0] != 0)
                {
                    for (int i = 0; i < field_node.Children.Count; i++)
                    {
                        TreeNode line_node = field_node.Children[i];
                        ProcessLineNode(line_node);
                    }
                }
            }
        }

        private void ProcessLineNode(TreeNode line_node, long search_level = 0)
        {
            var line_node_content = line_node.Content;
            if (line_node_content.NodeLabel == "Terminal")
            {
                ProcessTerminalNode(line_node, search_level);
                return;
            }

            int line_index = 0;
            while (line_index < line_node_content.Lines.Count)
            {
                long line_number = line_node_content.Lines[line_index];
                bool check_status = false;
                long found_paragraph_index = 0;
                if (line_mapping.ContainsKey(line_number))
                {
                    found_paragraph_index = line_node_content.HorizontalParagraph;
                    Regex regex_obj = Utilities.CreateRegexpObject(line_node_content.RE_Pattern);
                    var line_checker = new LineContentChecker(line_mapping[line_number], found_paragraph_index);
                    check_status = line_checker.CheckLineContents(regex_obj, line_node_content.CheckValue);
                    line_node_content.FoundValue = line_checker.joined_matches;
                    found_paragraph_index = line_checker.paragraph_index;
                }
                if (check_status)
                {
                    SetOffsetParagraph(line_node, found_paragraph_index, search_level);
                    int child_index = 0;
                    while (child_index < line_node.Children.Count)
                    {
                        TreeNode child_node = line_node.Children[child_index];
                        ProcessLineNode(child_node, search_level + 1);
                        child_index++;
                    }
                }
                else
                {
                    OffsetSearch(line_number, line_node, search_level, true);
                }
                line_index++;
            }
        }

        private void ProcessTerminalNode(TreeNode line_node, long search_level)
        {
            ProcessValue(line_node, search_level);
        }

        private void ProcessValue(TreeNode node, long search_level)
        {
            var node_content = node.Content;
            var parent_node = node.Parent;
            foreach (long line_number in node_content.Lines)
            {
                if (!line_mapping.ContainsKey(line_number))
                {
                    OffsetSearch(line_number, node, search_level, true);
                }
                else
                {
                    List<ParagraphContainer> paragraph_collection = line_mapping[line_number];
                    int start_index = node_content.HorizontalParagraph >= paragraph_collection.Count ? 0 : (int)node_content.HorizontalParagraph;

                    for (int paragraph_index = start_index; paragraph_index < paragraph_collection.Count; paragraph_index++)
                    {
                        string paragraph_text = paragraph_collection[paragraph_index].Text;
                        Regex regex_obj = Utilities.CreateRegexpObject(node_content.RE_Pattern);

                        MatchProcessor match_processor = new MatchProcessor(paragraph_text, regex_obj, node_content.ValueType);
                        if (!string.IsNullOrEmpty(match_processor.result))
                        {
                            node_content.FoundValue = match_processor.result;
                            node_content.Status = true;
                            PropagateStatusInTree(true, node);
                            return;
                        }
                    }
                }
            }
        }

        private void PropagateStatusInTree(bool status, TreeNode node)
        {
            TreeNode temp_node = node;
            while (temp_node.Parent.Content.Name != "root")
            {
                temp_node = temp_node.Parent;
                temp_node.Content.Status = status;
            }
        }

        private void SetOffsetChildrenParagraphs(TreeNode node, long search_level)
        {
            var node_content = node.Content;
            var field_config = (Dictionary<string, object>)config_data[node_content.Name];
            var search_parameters = (ArrayList)field_config["values"];
            var single_parameter = (Dictionary<string, dynamic>)search_parameters[(int)search_level];
            long horizontal_offset = single_parameter["horizontal_offset"];

            foreach (TreeNode child in node.Children)
            {
                var child_content = child.Content;
                child_content.HorizontalParagraph = node_content.HorizontalParagraph + horizontal_offset;
            }
        }

        private void SetOffsetParagraph(TreeNode line_node, long found_paragraph_index, long search_level)
        {
            line_node.Content.HorizontalParagraph = found_paragraph_index;
            SetOffsetChildrenParagraphs(line_node, search_level);
        }

        private void UpdateFieldNode(Dictionary<string, SimilarityDescription> collected_data, TreeNode field_node)
        {
            field_node.Children = new List<TreeNode>();
            if (field_node.Content.Lines.Count != 0)
            {
                field_node.Content.Lines.RemoveAt(0);
            }
            double max_similarity = collected_data.Values.ToList().Max(item => item.Ratio);
            AddChildrenToFieldNode(field_node, collected_data, max_similarity);
            tree_structure.AddSearchValues(config_data[field_node.Content.Name], field_node);
        }
    }
}