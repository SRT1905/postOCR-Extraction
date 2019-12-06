using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    internal class WordParser
    {
        private const long similarity_search_threshold = 8;
        private static SearchTree tree_structure;
        private readonly SortedDictionary<long, List<ParagraphContainer>> line_mapping;
        private Dictionary<string, object> config_data;

        public WordParser(SortedDictionary<long, List<ParagraphContainer>> document_content)
        {
            line_mapping = document_content;
        }

        public Dictionary<string, string> ParseDocument(Dictionary<string, object> config_data)
        {
            this.config_data = config_data;
            tree_structure = new SearchTree(config_data);
            ProcessDocument();
            return tree_structure.GetValuesFromTree();
        }

        private void AddChildrenToFieldNode(ref TreeNode field_node, Dictionary<string, SimilarityDescription> collected_data, double max_similarity)
        {
            foreach (string key in collected_data.Keys)
            {
                if (max_similarity == collected_data[key].Ratio)
                {
                    AddSingleChildToFieldNode(ref field_node, key);
                }
            }
        }

        private void AddSingleChildToFieldNode(ref TreeNode field_node, string key)
        {
            long line = long.Parse(key.Split('|')[0]);
            if (!field_node.Lines.Contains(line))
            {
                field_node.Lines.Add(line);
                field_node.AddChild(found_line: line, pattern: field_node.RE_Pattern, new_value: field_node.Value, node_label: "Line");
            }
        }

        private bool CheckLineContents(long line_number, Regex regex_obj, string search_value, ref long paragraph_index)
        {
            List<ParagraphContainer> paragraph_collection = line_mapping[line_number];
            long start_index = paragraph_index == 0 || paragraph_index >= paragraph_collection.Count ? 0 : paragraph_index;

            for (paragraph_index = start_index; paragraph_index < paragraph_collection.Count; paragraph_index++)
            {
                string paragraph_text = paragraph_collection[(int)paragraph_index].Text;
                if (regex_obj.IsMatch(paragraph_text))
                {
                    if (!string.IsNullOrEmpty(search_value))
                    {
                        if (GetMatchesFromParagraph(paragraph_text, regex_obj, search_value).Count != 0)
                        {
                            return true;
                        }
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            paragraph_index = 0;
            return false;
        }

        private void GetDataFromUndefinedNode(ref TreeNode field_node)
        {
            Regex regex_object = Utilities.CreateRegexpObject(field_node.RE_Pattern);
            Dictionary<string, SimilarityDescription> collected_data = new Dictionary<string, SimilarityDescription>();

            foreach (long line in line_mapping.Keys)
            {
                foreach (ParagraphContainer container in line_mapping[line])
                {
                    if (regex_object.IsMatch(container.Text))
                    {
                        List<SimilarityDescription> matched_data_collection = GetMatchesFromParagraph(container.Text, regex_object, field_node.Value);

                        for (int i = 0; i < matched_data_collection.Count; i++)
                        {
                            collected_data.Add($"{line}|{line_mapping[line].IndexOf(container)}|{i}", matched_data_collection[i]);
                        }
                    }
                }
            }

            if (collected_data.Count != 0)
            {
                UpdateFieldNode(collected_data, ref field_node);
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

        private List<long> GetOffsetLines(long line_number, string pattern, string search_value)
        {
            Regex regex_obj = Utilities.CreateRegexpObject(pattern);
            List<long> found_values_collection = new List<long>();

            for (int search_offset = 1; search_offset <= similarity_search_threshold; search_offset++)
            {
                List<long> offset_indexes = new List<long>() { line_number + search_offset, line_number - search_offset };
                foreach (long offset_index in offset_indexes)
                {
                    if (line_mapping.ContainsKey(offset_index))
                    {
                        long paragraph_index = 0;
                        if (CheckLineContents(offset_index, regex_obj, search_value, ref paragraph_index))
                        {
                            found_values_collection.Add(offset_index);
                        }
                    }
                }
            }
            return found_values_collection;
        }

        private void OffsetSearch(long line_number, ref TreeNode line_node, ref TreeNode parent_node, long search_level, bool add_to_parent = false)
        {
            List<long> line_numbers = GetOffsetLines(line_number, line_node.RE_Pattern, line_node.Value);

            foreach (long offset_index in line_numbers)
            {
                if (add_to_parent)
                {
                    if (!parent_node.Lines.Contains(offset_index))
                    {
                        parent_node.Lines.Add(offset_index);
                        TreeNode new_node = parent_node.AddChild(line_node.Name, line_node.RE_Pattern, offset_index, line_node.ValueType, line_node.Value, line_node.NodeLabel, line_node.HorizontalParagraph);
                        tree_structure.AddSearchValues(config_data[line_node.Name], ref new_node, (int)search_level);
                    }
                }
                else
                {
                    if (!line_node.Lines.Contains(offset_index))
                    {
                        line_node.Lines.Add(offset_index);
                        TreeNode new_node = line_node.AddChild(found_line: offset_index, pattern: line_node.RE_Pattern);
                        tree_structure.AddSearchValues(config_data[line_node.Name], ref new_node, (int)search_level);
                    }
                }
            }
        }

        private void ProcessDocument()
        {
            for (int field_index = 0; field_index < tree_structure.Children.Count; field_index++)
            {
                TreeNode field_node = tree_structure.Children[field_index];
                if (field_node.Lines[0] == 0)
                {
                    GetDataFromUndefinedNode(ref field_node);
                }
                if (field_node.Lines[0] != 0)
                {
                    for (int i = 0; i < field_node.Children.Count; i++)
                    {
                        TreeNode line_node = field_node.Children[i];
                        ProcessLineNode(ref line_node);
                    }
                }
            }
        }

        private void ProcessLineNode(ref TreeNode line_node, long search_level = 0)
        {
            if (line_node.NodeLabel == "Terminal")
            {
                ProcessTerminalNode(ref line_node, search_level);
                return;
            }

            int line_index = 0;
            while (line_index < line_node.Lines.Count)
            {
                long line_number = line_node.Lines[line_index];
                bool check_status = false;
                long found_paragraph_index = 0;
                if (line_mapping.ContainsKey(line_number))
                {
                    found_paragraph_index = line_node.HorizontalParagraph;
                    Regex regex_obj = Utilities.CreateRegexpObject(line_node.RE_Pattern);
                    check_status = CheckLineContents(line_number, regex_obj, line_node.Value, ref found_paragraph_index);
                }
                if (check_status)
                {
                    SetOffsetParagraph(ref line_node, found_paragraph_index, search_level);
                    int child_index = 0;
                    while (child_index < line_node.Children.Count)
                    {
                        TreeNode child_node = line_node.Children[child_index];
                        ProcessLineNode(ref child_node, search_level + 1);
                        child_index++;
                    }
                }
                else
                {
                    TreeNode parent_node = line_node.Parent;
                    OffsetSearch(line_number, ref line_node, ref parent_node, search_level, true);
                }
                line_index++;
            }
        }

        private void ProcessTerminalNode(ref TreeNode line_node, long search_level)
        {
            TreeNode parent_node = line_node.Parent;
            ProcessValue(ref line_node, ref parent_node, search_level);
        }

        private void ProcessValue(ref TreeNode node, ref TreeNode parent_node, long search_level)
        {
            foreach (long line_number in node.Lines)
            {
                if (!line_mapping.ContainsKey(line_number))
                {
                    OffsetSearch(line_number, ref node, ref parent_node, search_level, true);
                }
                else
                {
                    List<ParagraphContainer> paragraph_collection = line_mapping[line_number];
                    int start_index = node.HorizontalParagraph >= paragraph_collection.Count ? 0 : (int)node.HorizontalParagraph;

                    for (int paragraph_index = start_index; paragraph_index < paragraph_collection.Count; paragraph_index++)
                    {
                        string paragraph_text = paragraph_collection[paragraph_index].Text;
                        Regex regex_obj = Utilities.CreateRegexpObject(node.RE_Pattern);

                        MatchProcessor match_processor = new MatchProcessor(paragraph_text, regex_obj, node.ValueType);
                        if (!string.IsNullOrEmpty(match_processor.result))
                        {
                            node.Value = match_processor.result;
                            node.Status = true;
                            PropagateStatusInTree(true, ref node);
                            return;
                        }
                    }
                }
            }
        }

        private void PropagateStatusInTree(bool status, ref TreeNode node)
        {
            TreeNode temp_node = node;
            while (temp_node.Parent.Name != "root")
            {
                temp_node = temp_node.Parent;
                temp_node.Status = status;
            }
        }

        private void SetOffsetChildrenParagraphs(ref TreeNode node, long search_level)
        {
            var field_config = (Dictionary<string, object>)config_data[node.Name];
            var search_parameters = (ArrayList)field_config["values"];
            var single_parameter = (Dictionary<string, dynamic>)search_parameters[(int)search_level];
            long horizontal_offset = single_parameter["horizontal_offset"];

            foreach (TreeNode child in node.Children)
            {
                child.HorizontalParagraph = node.HorizontalParagraph + horizontal_offset;
            }
        }

        private void SetOffsetParagraph(ref TreeNode line_node, long found_paragraph_index, long search_level)
        {
            line_node.HorizontalParagraph = found_paragraph_index;
            SetOffsetChildrenParagraphs(ref line_node, search_level);
        }

        private void UpdateFieldNode(Dictionary<string, SimilarityDescription> collected_data, ref TreeNode field_node)
        {
            field_node.Children = new List<TreeNode>();
            field_node.Lines.RemoveAt(0);
            double max_similarity = collected_data.Values.ToList().Max(item => item.Ratio);
            AddChildrenToFieldNode(ref field_node, collected_data, max_similarity);
            tree_structure.AddSearchValues(config_data[field_node.Name], ref field_node);
        }
    }
}