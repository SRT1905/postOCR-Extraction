using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace SmartOCR
{
    class OutputWorkbook
    {
        private static Workbook instance;
        private static Worksheet output_worksheet;

        public static Workbook GetOutputWorkbook(string doc_type)
        {
            if (instance == null)
            {
                instance = CreateOutputWorkbook(doc_type);
            }
            return instance;
        }

        private static Workbook CreateOutputWorkbook(string doc_type)
        {
            if (ConfigParser.config_wb == null)
                new ConfigParser();
            Workbook source_wb = ConfigParser.config_wb;
            Workbook new_wb = ExcelApplication.GetExcelApplication().Workbooks.Add();

            Worksheet source_ws = GetWorksheetByName(source_wb, doc_type);
            if (source_ws == null)
                return null;

            output_worksheet = new_wb.Worksheets.Add(After: new_wb.Worksheets[new_wb.Worksheets.Count]);
            output_worksheet.Name = source_ws.Name;
            int i;
            for (i = 1; i <= source_ws.Cells[source_ws.Rows.Count, 1].End[XlDirection.xlUp].Row; i++)
            {
                if (source_ws.Cells[i, 1].Value2.ToString().ToLower().Contains("field name"))
                    break;
            }
            Range header_range = source_ws.Range[source_ws.Cells[i, 2], source_ws.Cells[i, source_ws.Columns.Count].End[XlDirection.xlToLeft]];
            Range first_cell = (Range)output_worksheet.Range[output_worksheet.Cells[1, 1],
                                                             output_worksheet.Cells[header_range.Rows.Count, header_range.Columns.Count]];
            first_cell.Value2 = header_range.Value2;
            first_cell.Rows.AutoFit();
            first_cell.Columns.AutoFit();

            new_wb.Worksheets[1].Delete();
            return new_wb;
        }

        private static Worksheet GetWorksheetByName(Workbook workbook, string sheet_name)
        {
            foreach (Worksheet item in workbook.Worksheets)
            {
                if (item.Name.Contains(sheet_name))
                    return item;
            }
            return null;
        }
        public static void ReturnValuesToWorksheet(Dictionary<string, string> values)
        {
            long row_to_input = GetLastRowInWorksheet();
            foreach (string key in values.Keys)
            {
                long column_index = FindColumnIndex(key);
                if (column_index != 0)
                    output_worksheet.Cells[row_to_input, column_index] = values[key];
            }
        }

        private static long FindColumnIndex(string field_name)
        {
            long last_column = output_worksheet.Cells[1, output_worksheet.Columns.Count].End[XlDirection.xlToLeft].Column;
            for (long i = 1; i <= last_column; i++)
            {
                if (output_worksheet.Cells[1, i].Value2 == field_name)
                    return i;
            }
            return 0;
        }

        private static long GetLastRowInWorksheet()
        {
            return output_worksheet.UsedRange.Rows[output_worksheet.UsedRange.Rows.Count].Row + 1;
        }
    }
}
