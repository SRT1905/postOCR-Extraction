using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// Used to create and populate output Excel workbook that would contain data collected from search tree.
    /// </summary>
    internal class ExcelOutputWorkbook
    {
        /// <summary>
        /// Output workbook instance.
        /// </summary>
        private static Workbook instance;
        /// <summary>
        /// Worksheet that would contain collected data.
        /// </summary>
        private static Worksheet output_worksheet;

        /// <summary>
        /// Gets output workbook.
        /// </summary>
        /// <returns></returns>
        public static Workbook GetOutputWorkbook()
        {
            if (instance == null)
            {
                instance = CreateOutputWorkbook();
            }
            return instance;
        }

        /// <summary>
        /// Creates new output workbook.
        /// </summary>
        /// <returns><see cref="Workbook"/> instance added to <see cref="Workbooks"/> collection.</returns>
        private static Workbook CreateOutputWorkbook()
        {
            if (ConfigParser.config_wb == null)
            {
                new ConfigParser();
            }

            Workbook source_wb = ConfigParser.config_wb;
            Workbook new_wb = ExcelApplication.GetExcelApplication().Workbooks.Add();

            Worksheet source_ws;
            try
            {
                source_ws = source_wb.Worksheets[1];
            }
            catch (Exception)
            {
                return null;
            }
            if (source_ws == null)
            {
                return null;
            }

            output_worksheet = new_wb.Worksheets.Add(After: new_wb.Worksheets[new_wb.Worksheets.Count]);
            output_worksheet.Name = source_ws.Name;
            int i;
            for (i = 1; i <= source_ws.Cells[source_ws.Rows.Count, 1].End[XlDirection.xlUp].Row; i++)
            {
                if (source_ws.Cells[i, 1].Value2.ToString().ToLower().Contains("field name"))
                {
                    break;
                }
            }
            Range header_range = source_ws.Range[source_ws.Cells[i, 2], source_ws.Cells[i, source_ws.Columns.Count].End[XlDirection.xlToLeft]];
            Range first_cell = (Range)output_worksheet.Cells.Item[1, 1];
            header_range.Copy(first_cell);

            new_wb.Worksheets[1].Delete();
            return new_wb;
        }

        /// <summary>
        /// Performs matching between worksheet field scheme and <paramref name="values"/> keys and returns values to sheet.
        /// </summary>
        /// <param name="values">Mapping between field names and found values.</param>
        public static void ReturnValuesToWorksheet(Dictionary<string, string> values)
        {
            long row_to_input = GetLastRowInWorksheet();
            var keys = values.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                string key = keys[i];
                long column_index = FindColumnIndex(key);
                if (column_index != 0)
                {
                    output_worksheet.Cells[row_to_input, column_index] = values[key];
                }
            }
        }

        /// <summary>
        /// Searches for match between <paramref name="field_name"/> and first row values.
        /// </summary>
        /// <param name="field_name">Field name to search.</param>
        /// <returns>Index of matched column.</returns>
        private static long FindColumnIndex(string field_name)
        {
            long last_column = output_worksheet.Cells[1, output_worksheet.Columns.Count].End[XlDirection.xlToLeft].Column;
            for (long i = 1; i <= last_column; i++)
            {
                if (output_worksheet.Cells[1, i].Value2 == field_name)
                {
                    return i;
                }
            }
            return 0;
        }

        /// <summary>
        /// Gets first empty row in <see cref="Worksheet"/> used range.
        /// </summary>
        /// <returns>Index of row.</returns>
        private static long GetLastRowInWorksheet()
        {
            return output_worksheet.UsedRange.Rows[output_worksheet.UsedRange.Rows.Count].Row + 1;
        }
    }
}