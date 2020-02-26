using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace SmartOCR
{
    public class WordParser // TODO: add summary.
    {
        #region Private constants        
        private const long similarity_search_threshold = 5;
        #endregion
        
        #region Fields
        private readonly ConfigData config_data;
        private readonly SortedDictionary<long, List<ParagraphContainer>> line_mapping;
        private readonly List<WordTable> tables;
        private readonly SearchTree tree_structure;
        #endregion

        #region Constructors
        public WordParser(WordReader reader, ConfigData configData)
        {
            if (reader == null)
            {
                throw new ArgumentNullException(nameof(reader));
            }
            line_mapping = reader.LineMapping;
            tables = reader.TableCollection;
            tree_structure = new SearchTree(configData);
            this.config_data = configData;
        }
        #endregion

        #region Public methods
        public Dictionary<string, string> ParseDocument()
        {
            tree_structure.PopulateTree();
            ProcessDocument();
            return tree_structure.GetValuesFromTree();
        }
        #endregion

        #region Private methods
        private void AddOffsetNode(TreeNode node, long search_level,
                           long offset_index, string found_value,
                           decimal position, bool add_to_parent)
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
                pattern = first_child_content.RegExPattern;
                horizontal_position = position;
            }
            else
            {
                node_label = node.Content.NodeLabel;
                pattern = node.Content.RegExPattern;
                horizontal_position = position;
            }
            TreeNode child_node = node.AddChild(foundLine: offset_index,
                                                pattern: pattern,
                                                nodeLabel: node_label,
                                                horizontalParagraph: horizontal_position);
            child_node.Content.FoundValue = found_value;
            SearchTree.AddSearchValues(config_data[node_content.Name], child_node, (int)search_level);
        }
        private void GetDataFromTableNode(TreeNode field_node)
        {
            TreeNodeContent content = field_node.Content;
            foreach (WordTable item in tables)
            {
                bool search_status = TryToFindMatchInTable(item, Utilities.CreateRegexpObject(content.RegExPattern),
                                                           content.CheckValue);
                if (search_status)
                {
                    System.Diagnostics.Debugger.Break();
                    TreeNode childNode = field_node.Children[0];
                    while (childNode.Content.NodeLabel != "Terminal")
                    {
                        childNode = childNode.Children[0];
                    }
                    TreeNodeContent childContent = childNode.Content;
                    string item_by_expression_position = item[childContent.FirstSearchParameter, childContent.SecondSearchParameter];
                    if (Utilities.CreateRegexpObject(childContent.RegExPattern).IsMatch(item_by_expression_position))
                    {
                        childContent.FoundValue = item_by_expression_position;
                        childContent.Status = true;
                        return;
                    }

                }
            }
        }
        private void GetDataFromUndefinedNode(TreeNode field_node)
        {
            Regex regex_object = Utilities.CreateRegexpObject(field_node.Content.RegExPattern);
            Dictionary<string, SimilarityDescription> collected_data = new Dictionary<string, SimilarityDescription>();

            var keys = line_mapping.Keys.ToList();
            for (int key_index = 0; key_index < keys.Count; key_index++)
            {
                long line = keys[key_index];
                for (int container_index = 0; container_index < line_mapping[line].Count; container_index++)
                {
                    ParagraphContainer container = line_mapping[line][container_index];
                    if (regex_object.IsMatch(container.Text))
                    {
                        var matched_data_collection = GetMatchesFromParagraph(container.Text,
                                                                              regex_object,
                                                                              field_node.Content.CheckValue);

                        for (int i = 0; i < matched_data_collection.Count; i++)
                        {
                            collected_data.Add($"{line}|{container_index}|{container.HorizontalLocation}|{i}", matched_data_collection[i]);
                        }
                    }
                }
            }

            if (collected_data.Count != 0)
            {
                UpdateFieldNode(collected_data, field_node);
            }
        }
        private Dictionary<string, string> GetOffsetLines(long line_number, TreeNodeContent content)
        {
            Regex regex_obj = Utilities.CreateRegexpObject(content.RegExPattern);
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
                            found_values_collection.Add(
                                string.Join("|",
                                            line,
                                            line_checker.ParagraphHorizontalLocation),
                                line_checker.JoinedMatches);
                        }
                    }
                }
            }
            return found_values_collection;
        }
        private void OffsetSearch(long line_number, TreeNode line_node,
                                  long search_level, bool add_to_parent = false)
        {
            TreeNodeContent line_node_content = (TreeNodeContent)line_node.Content;
            var line_numbers = GetOffsetLines(line_number, line_node_content);

            var keys = line_numbers.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                string[] splitted_key = key.Split('|');
                long offset_index = long.Parse(splitted_key[0],
                                               NumberStyles.Any,
                                               NumberFormatInfo.CurrentInfo);
                decimal horizontal_position = decimal.Parse(splitted_key[1],
                                                            NumberStyles.Any,
                                                            NumberFormatInfo.CurrentInfo);
                if (add_to_parent)
                {
                    var parent = line_node.Parent;
                    AddOffsetNode(parent, search_level,
                                  offset_index, line_numbers[key],
                                  horizontal_position, add_to_parent);
                }
                else
                {
                    AddOffsetNode(line_node, search_level,
                                  offset_index, line_numbers[key],
                                  horizontal_position, add_to_parent);
                }
            }
        }
        private void ProcessDocument()
        {
            for (int field_index = 0; field_index < tree_structure.Children.Count; field_index++)
            {
                TreeNode field_node = tree_structure.Children[field_index];
                TreeNodeContent node_content = field_node.Content;

                if (!node_content.ValueType.Contains("Table"))
                {
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
                else
                {
                    GetDataFromTableNode(field_node);
                }
            }
        }
        private void ProcessLineNode(TreeNode line_node, long search_level = 0)
        {
            TreeNodeContent line_node_content = (TreeNodeContent)line_node.Content;
            if (line_node_content.NodeLabel == "Terminal")
            {
                ProcessValue(line_node, search_level);
                return;
            }

            int line_index = 0;
            while (line_index < line_node_content.Lines.Count)
            {
                long line_number = line_node_content.Lines[line_index];

                bool check_status = line_mapping.ContainsKey(line_number)
                    ? TryMatchLineData(line_node_content, line_number)
                    : false;
                if (check_status)
                {
                    SetOffsetChildrenLines(line_node, line_number);
                    ProcessLineNodeChildren(line_node, search_level);
                }
                else
                {
                    OffsetSearch(line_number, line_node, search_level, true);
                }
                line_index++;
            }
        }
        private void ProcessLineNodeChildren(TreeNode line_node, long search_level)
        {
            int child_index = 0;
            while (child_index < line_node.Children.Count)
            {
                TreeNode child_node = line_node.Children[child_index];
                ProcessLineNode(child_node, search_level + 1);
                child_index++;
            }
        }
        private void ProcessValue(TreeNode node, long search_level)
        {
            TreeNodeContent node_content = (TreeNodeContent)node.Content;
            for (int i = 0; i < node_content.Lines.Count; i++)
            {
                long line_number = node_content.Lines[i];
                if (!line_mapping.ContainsKey(line_number))
                {
                    OffsetSearch(line_number, node, search_level, true);
                }
                else
                {
                    List<ParagraphContainer> paragraph_collection = line_mapping[line_number];
                    int start_index = 0;
                    int finish_index = paragraph_collection.Count - 1;
                    switch (node_content.SecondSearchParameter)
                    {
                        case 1:
                            start_index = GetParagraphByLocation(paragraph_collection,
                                                                 node_content.HorizontalParagraph,
                                                                 return_next_largest: true);
                            finish_index = paragraph_collection.Count - 1;
                            break;
                        case -1:
                            start_index = 0;
                            finish_index = GetParagraphByLocation(paragraph_collection,
                                                                  node_content.HorizontalParagraph,
                                                                  return_next_largest: false);
                            break;
                        default:
                            break;
                    }

                    for (int paragraph_index = start_index; paragraph_index <= finish_index; paragraph_index++)
                    {
                        string paragraph_text = paragraph_collection[paragraph_index].Text;
                        Regex regex_obj = Utilities.CreateRegexpObject(node_content.RegExPattern);

                        MatchProcessor match_processor = new MatchProcessor(paragraph_text,
                                                                            regex_obj,
                                                                            node_content.ValueType);
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
        private bool TryMatchLineData(TreeNodeContent line_node_content, long line_number)
        {
            bool check_status;
            decimal paragraph_horizontal_location = line_node_content.HorizontalParagraph;
            Regex regex_obj = Utilities.CreateRegexpObject(line_node_content.RegExPattern);
            var line_checker = new LineContentChecker(line_mapping[line_number],
                                                      paragraph_horizontal_location,
                                                      line_node_content.SecondSearchParameter);
            check_status = line_checker.CheckLineContents(regex_obj, line_node_content.CheckValue);
            if (check_status)
            {
                line_node_content.HorizontalParagraph = line_checker.ParagraphHorizontalLocation;
            }
            else
            {
                line_node_content.HorizontalParagraph = paragraph_horizontal_location;
            }
            line_node_content.FoundValue = line_checker.JoinedMatches;
            return check_status;
        }
        private void SetOffsetChildrenLines(TreeNode node, long line)
        {
            TreeNodeContent node_content = (TreeNodeContent)node.Content;

            for (int i = 0; i < node.Children.Count; i++)
            {
                TreeNode child = node.Children[i];
                TreeNodeContent child_content = (TreeNodeContent)child.Content;
                child_content.HorizontalParagraph = node_content.HorizontalParagraph;
                List<long> keys = line_mapping.Keys.ToList();
                int line_index = keys.IndexOf(line) + child_content.FirstSearchParameter;
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
            field_node.Children.Clear();
            if (field_node.Content.Lines.Count != 0)
            {
                field_node.Content.Lines.RemoveAt(0);
            }
            double max_similarity = collected_data.Values.ToList().Max(item => item.Ratio);
            AddChildrenToFieldNode(field_node, collected_data, max_similarity);
            SearchTree.AddSearchValues(config_data[field_node.Content.Name], field_node);
        }
        #endregion

        #region Private static methods
        private static void AddChildrenToFieldNode(TreeNode field_node,
                                                   Dictionary<string, SimilarityDescription> collected_data,
                                                   double max_similarity)
        {
            var keys = collected_data.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                if (max_similarity == collected_data[key].Ratio)
                {
                    AddSingleChildToFieldNode(field_node, key);
                }
            }
        }
        private static void AddSingleChildToFieldNode(TreeNode field_node, string key)
        {
            var content = field_node.Content;
            long line = long.Parse(key.Split('|')[0],
                                   NumberStyles.Any,
                                   NumberFormatInfo.CurrentInfo);
            decimal horizontal_location = decimal.Parse(key.Split('|')[2],
                                                        NumberStyles.Any,
                                                        NumberFormatInfo.CurrentInfo);
            if (!content.Lines.Contains(line))
            {
                content.Lines.Add(line);
                field_node.AddChild(foundLine: line, pattern: content.RegExPattern,
                                    newValue: content.CheckValue, nodeLabel: "Line",
                                    horizontalParagraph: horizontal_location);
            }
        }
        private static List<SimilarityDescription> GetMatchesFromParagraph(string text_to_check,
                                                                           Regex re_object,
                                                                           string check_value)
        {
            MatchCollection matches = re_object.Matches(text_to_check);
            List<SimilarityDescription> found_values = new List<SimilarityDescription>();
            for (int i = 0; i < matches.Count; i++)
            {
                Match single_match = matches[i];
                if (single_match.Groups.Count > 1)
                {
                    for (int group_i = 1; group_i < single_match.Groups.Count; group_i++)
                    {
                        Group group_item = single_match.Groups[group_i];
                        SimilarityDescription description = new SimilarityDescription(group_item.Value,
                                                                                      check_value);
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
        private static int GetParagraphByLocation(List<ParagraphContainer> paragraph_collection,
                                                  decimal position,
                                                  bool return_next_largest)
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
        private static void PropagateStatusInTree(bool status, TreeNode node)
        {
            TreeNode temp_node = node;
            while (temp_node.Parent.Content.Name != "root")
            {
                temp_node = temp_node.Parent;
                temp_node.Content.Status = status;
            }
        }
        private static bool TryToFindMatchInTable(WordTable table, Regex regex_obj, string check_value)
        {
            for (int i = 0; i < table.RowCount; i++)
            {
                for (int j = 0; j < table.ColumnCount; j++)
                {
                    if (table[i, j] != null && regex_obj.IsMatch(table[i, j]))
                    {
                        Match single_match = regex_obj.Match(table[i, j]);
                        SimilarityDescription similarity;
                        if (single_match.Groups.Count > 0)
                        {
                            similarity = new SimilarityDescription(single_match.Groups[1].Value, check_value);
                        }
                        else
                        {
                            similarity = new SimilarityDescription(single_match.Value, check_value);
                        }
                        if (similarity.CheckStringSimilarity())
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        #endregion
    }
}