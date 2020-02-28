namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Used to create and populate output Excel workbook that would contain data collected from search tree.
    /// </summary>
    public static class ExcelOutputWorkbook
    {
        /// <summary>
        /// Output workbook instance.
        /// </summary>
        private static Workbook instance;

        /// <summary>
        /// Worksheet that would contain collected data.
        /// </summary>
        private static Worksheet outputWorksheet;

        /// <summary>
        /// Gets output workbook.
        /// </summary>
        /// <returns>An instance of <see cref="Workbook"/> object.</returns>
        public static Workbook GetOutputWorkbook()
        {
            if (instance == null)
            {
                instance = CreateOutputWorkbook();
            }

            return instance;
        }

        /// <summary>
        /// Performs matching between worksheet field scheme and <paramref name="values"/> keys and returns values to sheet.
        /// </summary>
        /// <param name="values">Mapping between field names and found values.</param>
        public static void ReturnValuesToWorksheet(Dictionary<string, string> values)
        {
            long rowToInput = GetLastRowInWorksheet();
            foreach (var item in values)
            {
                ProcessSingleValue(rowToInput, item);
            }
        }

        private static void ProcessSingleValue(long rowToInput, KeyValuePair<string, string> item)
        {
            long columnIndex = FindColumnIndex(item.Key);
            if (columnIndex != 0)
            {
                outputWorksheet.Cells[rowToInput, columnIndex] = item.Value;
            }
        }

        /// <summary>
        /// Creates new output workbook.
        /// </summary>
        /// <returns><see cref="Workbook"/> instance added to <see cref="Workbooks"/> collection.</returns>
        private static Workbook CreateOutputWorkbook()
        {
            Workbook sourceWB = ConfigParser.ConfigWorkbook;
            Workbook newWB = ExcelApplication.AddEmptyWorkbook();

            Worksheet sourceWS;
            try
            {
                sourceWS = sourceWB.Worksheets[1];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }

            if (sourceWS == null)
            {
                return null;
            }

            outputWorksheet = newWB.Worksheets.Add(After: newWB.Worksheets[newWB.Worksheets.Count]);
            outputWorksheet.Name = sourceWS.Name;
            int i;
            for (i = 1; i <= sourceWS.Cells[sourceWS.Rows.Count, 1].End[XlDirection.xlUp].Row; i++)
            {
                if (sourceWS.Cells[i, 1].Value2.ToString().ToLower().Contains("field name"))
                {
                    break;
                }
            }

            CopyHeaderBetweenWorkbooks(newWB, sourceWS, i);
            return newWB;
        }

        private static void CopyHeaderBetweenWorkbooks(Workbook newWB, Worksheet sourceWS, int rowIndex)
        {
            Range headerRange = sourceWS.Range[sourceWS.Cells[rowIndex, 2], sourceWS.Cells[rowIndex, sourceWS.Columns.Count].End[XlDirection.xlToLeft]];
            headerRange.Copy((Range)outputWorksheet.Cells.Item[1, 1]);

            newWB.Worksheets[1].Delete();
        }

        /// <summary>
        /// Searches for match between <paramref name="fieldName"/> and first row values.
        /// </summary>
        /// <param name="fieldName">Field name to search.</param>
        /// <returns>Index of matched column.</returns>
        private static long FindColumnIndex(string fieldName)
        {
            long lastColumn = outputWorksheet.UsedRange.Columns.Count; // .Cells[1, outputWorksheet.Columns.Count].End[XlDirection.xlToLeft].Column;
            for (long i = 1; i <= lastColumn; i++)
            {
                if (outputWorksheet.Cells[1, i].Value2 == fieldName)
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
            return outputWorksheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row + 1; // .Rows[outputWorksheet.UsedRange.Rows.Count].Row + 1;
        }
    }
}