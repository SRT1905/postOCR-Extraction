using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace SmartOCR
{
    class ConfigParser
    {
        public static Workbook config_wb;

        public ConfigParser()
        {
            if (config_wb == null)
                config_wb = GetInternalConfig();
        }

        private Workbook GetInternalConfig()
        {
            string temp_path = Path.GetTempFileName();
            File.WriteAllBytes(temp_path, ConfigContainer.config);
            return ExcelApplication.OpenExcelWorkbook(temp_path);
        }

        public Dictionary<string, object> ParseConfig(string doc_type)
        {
            foreach (Worksheet item in config_wb.Worksheets)
            {
                if (item.Name.Contains(doc_type))
                {
                    return GetFields(item);
                }
            }
            return null;
        }
        private Dictionary<string, object> GetFields(Worksheet source_ws)
        {
            Dictionary<string, object> fields = new Dictionary<string, object>();
            long header_row = 1;
            string cell_value;
            for (int i = 1; i <= source_ws.Cells.Item[source_ws.Rows.Count, 1].End[XlDirection.xlUp].Row; i++)
            {
                cell_value = source_ws.Cells.Item[i, 1].Value2;
                if (cell_value.ToLower().Contains("field name"))
                {
                    header_row = i;
                    break;
                }
            }
            for (int i = 2; i <= source_ws.Cells.Item[header_row, source_ws.Columns.Count].End[XlDirection.xlToLeft].Column; i++)
            {
                cell_value = source_ws.Cells.Item[header_row, i].Value2;
                fields.Add(cell_value, GetSearchParameters(source_ws, cell_value));
            }
            return fields;
        }

        private Dictionary<string, object> GetSearchParameters(Worksheet source_ws, string field_name)
        {
            long item_column;
            for (item_column = 2; item_column <= source_ws.UsedRange.Columns.Count; item_column++)
                if (source_ws.Cells.Item[1, item_column].Value2 == field_name)
                    break;
            if (source_ws.Cells.Item[1, item_column].Value2 != field_name)
            {
                Console.WriteLine($"Processing error - field name '{field_name}' not found.");
                return null;
            }
            long type_row = 0;
            long field_expression_row = 0;
            long value_row = 0;
            long last_row = source_ws.UsedRange.Rows.Count;
            long i;
            for (i = 1; i <= last_row; i++)
            {
                string cell_value = source_ws.Cells.Item[i, 1].Value2;
                if (cell_value == null)
                    cell_value = string.Empty;
                if (cell_value.ToLower().Contains("value type"))
                    type_row = i;
                if (cell_value.ToLower().Contains("field expression"))
                    field_expression_row = i;
                if (cell_value.ToLower().Contains("search values"))
                    value_row = i;
                if (type_row > 0 && value_row > 0 &&
                    field_expression_row > 0 && string.IsNullOrEmpty(source_ws.Cells.Item[i, item_column].Value2))
                {
                    break;
                }
            }
            return new Dictionary<string, object>
            {
                { "type", source_ws.Cells.Item[type_row, item_column].Value2 },
                { "field_expression", GetFieldExpression(source_ws.Cells.Item[field_expression_row, item_column].Value2, field_name) },
                { "paragraphs", new List<long>(){ 0} },
                { "values", GetPatterns(source_ws, item_column, value_row, i - 1) }
            };
        }

        private Dictionary<string, object> GetFieldExpression(string cell_value, string field_name)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            string[] splitted_value = cell_value.Split(';');
            if (string.IsNullOrEmpty(splitted_value[0]))
                parameters.Add("regexp", @field_name);
            else
                parameters.Add("regexp", @splitted_value[0]);
            if (string.IsNullOrEmpty(splitted_value[1]))
                parameters.Add("value", field_name);
            else
                parameters.Add("value", splitted_value[1]);
            return parameters;
        }

        private ArrayList GetPatterns(Worksheet source_ws, long item_column, long first_row, long last_row)
        {
            ArrayList patterns = new ArrayList();
            for (long i = first_row; i <= last_row; i++)
            {
                ArrayList splitted_value = GetPatternDefinition(source_ws, i, item_column);
                Dictionary<string, object> search_value_container = new Dictionary<string, object>
                {
                    {"regexp", @splitted_value[0] },
                    {"offset", splitted_value[1] },
                    {"horizontal_offset", splitted_value[2] }
                };
                patterns.Add(search_value_container);
            }
            return patterns;
        }

        private ArrayList GetPatternDefinition(Worksheet source_ws, long item_row, long item_column)
        {
            string cell_value = source_ws.Cells.Item[item_row, item_column].Value2;

            if (string.IsNullOrEmpty(cell_value))
                return null;
            ArrayList values = new ArrayList(cell_value.Split(';'));
            while (values.Count < 3)
            {
                values.Add(string.Empty);
            }
            ArrayList splitted_offset = new ArrayList(values[1].ToString().Split(','));
            while (splitted_offset.Count < 2)
            {
                splitted_offset.Add(string.Empty);
            }

            if (string.IsNullOrEmpty(values[1].ToString()))
            {
                values[1] = 0;
                values[2] = 0;
            }
            else
            {
                values[1] = string.IsNullOrEmpty(splitted_offset[0].ToString()) ? 0 : long.Parse(splitted_offset[0].ToString());
                if (splitted_offset[1] == null)
                    values[2] = 0;
                else
                    values[2] = string.IsNullOrEmpty(splitted_offset[1].ToString()) ? 0 : long.Parse(splitted_offset[1].ToString());
            }
            return values;
        }
    }
}
