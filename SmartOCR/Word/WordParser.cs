using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    internal class WordParser // TODO: add summary.
    {
        private const long similarity_search_threshold = 8;

        private readonly SearchTree tree_structure;
        private readonly SortedDictionary<long, List<ParagraphContainer>> line_mapping;
        private readonly ConfigData config_data;

        public WordParser(SortedDictionary<long, List<ParagraphContainer>> document_content, ConfigData config_data)
        {
            line_mapping = document_content;
            tree_structure = new SearchTree(config_data);
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
            decimal horizontal_location = decimal.Parse(key.Split('|')[2]);
            if (!content.Lines.Contains(line))
            {
                content.Lines.Add(line);
                field_node.AddChild(found_line: line, pattern: content.RE_Pattern, new_value: content.CheckValue, node_label: "Line", horizontal_paragraph: horizontal_location);
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
                            collected_data.Add($"{line}|{line_mapping[line].IndexOf(container)}|{container.HorizontalLocation}|{i}", matched_data_collection[i]);
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

        private Dictionary<string, string> GetOffsetLines(long line_number, TreeNodeContent content)
        {
            Regex regex_obj = Utilities.CreateRegexpObject(content.RE_Pattern);
            var found_values_collection = new Dictionary<string, string>();
            List<long> keys = line_mapping.Keys.ToList();
            int line_index = keys.IndexOf(line_number);
            for (int search_offset = 1; search_offset <= similarity_search_threshold; search_offset++)
            {
                List<int> offset_indexes = new List<int>() { line_index + search_offset, line_index - search_offset };
                foreach (int offset_index in offset_indexes)
                {
                    if (offset_index >= 0 && offset_index < line_mapping.Count)
                    {
                        long line = keys[offset_index];
                        var line_checker = new LineContentChecker(line_mapping[line]);
                        if (line_checker.CheckLineContents(regex_obj, content.CheckValue))
                        {
                            found_values_collection.Add(string.Join("|", line, line_checker.paragraph_horizontal_location), line_checker.joined_matches);
                        }
                    }
                }
            }
            return found_values_collection;
        }


        private void OffsetSearch(long line_number, TreeNode line_node, long search_level, bool add_to_parent = false)
        {
            var line_node_content = line_node.Content;
            var line_numbers = GetOffsetLines(line_number, line_node_content);

            foreach (string key in line_numbers.Keys)
            {
                string[] splitted_key = key.Split('|');
                long offset_index = long.Parse(splitted_key[0]);
                decimal horizontal_position = decimal.Parse(splitted_key[1]);
                if (add_to_parent)
                {
                    var parent = line_node.Parent;
                    AddOffsetNode(parent, search_level, offset_index, line_numbers[key], horizontal_position, add_to_parent);
                }
                else
                {
                    AddOffsetNode(line_node, search_level, offset_index, line_numbers[key], horizontal_position, add_to_parent);
                }
            }
        }

        private void AddOffsetNode(TreeNode node, long search_level, long offset_index, string found_value, decimal position, bool add_to_parent)
        {
            var node_content = node.Content;
            if (node_content.Lines.Count(item => item == offset_index) >= 2)
            {
                return;
            }
            node_content.Lines.Add(offset_index);
            string node_label;
            string pattern;
            decimal horizontal_position;
            if (add_to_parent)
            {
                var first_child_content = node.Children.First().Content;
                node_label = first_child_content.NodeLabel;
                pattern = first_child_content.RE_Pattern;
                horizontal_position = position;
            }
            else
            {
                node_label = node.Content.NodeLabel;
                pattern = node.Content.RE_Pattern;
                horizontal_position = position;
            }
            TreeNode child_node = node.AddChild(found_line: offset_index, pattern: pattern, node_label: node_label, horizontal_paragraph: horizontal_position) ;
            child_node.Content.FoundValue = found_value;
            tree_structure.AddSearchValues(config_data[node_content.Name], child_node, (int)search_level);
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
                ProcessValue(line_node, search_level);
                return;
            }

            int line_index = 0;
            while (line_index < line_node_content.Lines.Count)
            {
                long line_number = line_node_content.Lines[line_index];
                bool check_status = false;
                if (line_mapping.ContainsKey(line_number))
                {
                    decimal paragraph_horizontal_location = line_node_content.HorizontalParagraph;
                    Regex regex_obj = Utilities.CreateRegexpObject(line_node_content.RE_Pattern);
                    var line_checker = new LineContentChecker(line_mapping[line_number], paragraph_horizontal_location, line_node_content.HorizontalStatus);
                    check_status = line_checker.CheckLineContents(regex_obj, line_node_content.CheckValue);
                    if (check_status)
                    {
                        line_node_content.HorizontalParagraph = line_checker.paragraph_horizontal_location;
                    }
                    else
                    {
                        line_node_content.HorizontalParagraph = paragraph_horizontal_location;
                    }
                    line_node_content.FoundValue = line_checker.joined_matches;
                }
                if (check_status)
                {
                    SetOffsetChildrenLines(line_node, line_number);
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

        private void ProcessValue(TreeNode node, long search_level)
        {
            var node_content = node.Content;
            foreach (long line_number in node_content.Lines)
            {
                if (!line_mapping.ContainsKey(line_number))
                {
                    OffsetSearch(line_number, node, search_level, true);
                }
                else
                {
                    List<ParagraphContainer> paragraph_collection = line_mapping[line_number];
                    int start_index = 0;
                    int finish_index = paragraph_collection.Count - 1;
                    switch (node_content.HorizontalStatus)
                    {
                        case 1:
                            start_index = GetParagraphByLocation(paragraph_collection, node_content.HorizontalParagraph, true);
                            finish_index = paragraph_collection.Count - 1;
                            break;
                        case -1:
                            start_index = 0;
                            finish_index = GetParagraphByLocation(paragraph_collection, node_content.HorizontalParagraph, false);
                            break;
                        default:
                            break;
                    }

                    for (int paragraph_index = start_index; paragraph_index <= finish_index; paragraph_index++)
                    {
                        string paragraph_text = paragraph_collection[paragraph_index].Text;
                        Regex regex_obj = Utilities.CreateRegexpObject(node_content.RE_Pattern);

                        MatchProcessor match_processor = new MatchProcessor(paragraph_text, regex_obj, node_content.ValueType);
                        if (!string.IsNullOrEmpty(match_processor.Result))
                        {
                            node_content.FoundValue = match_processor.Result;
                            node_content.Status = true;
                            PropagateStatusInTree(true, node);
                            return;
                        }
                    }

                    OffsetSearch(line_number, node, search_level, true);
                }
            }
        }

        private int GetParagraphByLocation(List<ParagraphContainer> paragraph_collection, decimal position, bool return_next_largest)
        {
            List<decimal> locations = paragraph_collection.Select(item => item.HorizontalLocation).ToList();
            int location = locations.BinarySearch(position);
            if (location < 0)
            {
                location = ~location;
            }
            if (return_next_largest)
            {
                if (location == paragraph_collection.Count)
                {
                    return location--;
                }
                return location;
            }
            return location--;
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

        private void SetOffsetChildrenLines(TreeNode node, long line)
        {
            var node_content = node.Content;

            foreach (TreeNode child in node.Children)
            {
                var child_content = child.Content;
                child_content.HorizontalParagraph = node_content.HorizontalParagraph;
                List<long> keys = line_mapping.Keys.ToList();
                int line_index = keys.IndexOf(line) + child_content.LineOffset;
                if (line_index >= 0 && line_index < keys.Count)
                {
                    long offset_line = keys[line_index];
                    child_content.Lines.Clear();
                    child_content.Lines.Add(offset_line);
                }
                
            }
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