using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace SmartOCR
{

    internal class ConfigParser // TODO: add summary.
    {
        public static Workbook config_wb;

        public ConfigParser()
        {
            if (config_wb == null)
            {
                config_wb = GetInternalConfigWorkbook();
            }
        }

        public ConfigParser(string config_file)
        {
            config_wb = GetExternalConfigWorkbook(config_file);
        }

        private Workbook GetExternalConfigWorkbook(string path)
        {
            return ExcelApplication.OpenExcelWorkbook(path);
        }

        private Workbook GetInternalConfigWorkbook()
        {
            string temp_path = Path.GetTempFileName();
            File.WriteAllBytes(temp_path, ConfigContainer.config);
            return ExcelApplication.OpenExcelWorkbook(temp_path);
        }
        
        public ConfigData ParseConfig()
        {
            return GetConfigData(config_wb.Worksheets[1]);
        }

        private ConfigData GetConfigData(Worksheet source_ws)
        {
            var data = new ConfigData();
            long header_row;
            for (header_row = 1; header_row <= source_ws.Cells.Item[source_ws.Rows.Count, 1].End[XlDirection.xlUp].Row; header_row++)
            {
                string cell_value = source_ws.Cells.Item[header_row, 1].Value2;
                if (cell_value.ToLower(CultureInfo.CurrentCulture).Contains("field name"))
                {
                    break;
                }
            }

            long last_column = source_ws.Cells.Item[header_row, source_ws.Columns.Count].End[XlDirection.xlToLeft].Column;
            for (int i = 2; i <= last_column; i++)
            {
                ConfigField field = GetFieldDefinition(source_ws, header_row, i);
                data.AddField(field);
            }
            return data;
        }

        private ConfigField GetFieldDefinition(Worksheet source_ws, long header_row, long field_column)
        {
            string field_name = source_ws.Cells.Item[header_row, field_column].Value2;
            string value_type = source_ws.Cells.Item[header_row + 1, field_column].Value2;
            var field = new ConfigField(field_name, value_type);
            field.ParseFieldExpression(source_ws.Cells.Item[header_row + 2, field_column].Value2);
            long last_row = source_ws.Cells.Item[source_ws.Rows.Count, field_column].End[XlDirection.xlUp].Row;
            for (long i = header_row + 3; i <= last_row; i++)
            {
                var expression = new ConfigExpression(source_ws.Cells.Item[i, field_column].Value2);
                field.AddSearchExpression(expression);
            }
            return field;
        }
    }
}