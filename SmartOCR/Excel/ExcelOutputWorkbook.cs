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
            Utilities.Debug($"Getting output workbook.");
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
            Utilities.Debug($"Inserting found data into Excel output workbook.", 1);
            int rowToInput = GetLastRowInWorksheet();
            foreach (var item in values)
            {
                ProcessSingleValue(rowToInput, item);
            }
        }

        private static void ProcessSingleValue(int rowToInput, KeyValuePair<string, string> item)
        {
            int columnIndex = FindColumnIndex(item.Key);
            if (columnIndex != 0)
            {
                outputWorksheet.Cells[rowToInput, columnIndex] = item.Value;
            }
        }

        private static Worksheet GetOutputWorksheet(Workbook source)
        {
            try
            {
                return source.Worksheets[1];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        /// <summary>
        /// Creates new output workbook.
        /// </summary>
        /// <returns><see cref="Workbook"/> instance added to <see cref="Workbooks"/> collection.</returns>
        private static Workbook CreateOutputWorkbook()
        {
            Worksheet sourceWS = GetOutputWorksheet(ConfigParser.ConfigWorkbook);

            return sourceWS == null
                ? null
                : DefineOutputWorksheet(ExcelApplication.AddEmptyWorkbook(), sourceWS);
        }

        private static Workbook DefineOutputWorksheet(Workbook newWB, Worksheet sourceWS)
        {
            InitializeOutputWorksheet(newWB, sourceWS);
            CopyHeaderBetweenWorkbooks(newWB, sourceWS);
            return newWB;
        }

        private static void InitializeOutputWorksheet(Workbook newWB, Worksheet sourceWS)
        {
            outputWorksheet = newWB.Worksheets.Add(After: newWB.Worksheets[newWB.Worksheets.Count]);
            outputWorksheet.Name = sourceWS.Name;
        }

        private static int GetIdentifyingRow(Worksheet sourceWS)
        {
            for (int i = 1; i <= sourceWS.Cells[sourceWS.Rows.Count, 1].End[XlDirection.xlUp].Row; i++)
            {
                if (sourceWS.Cells[i, 1].Value2
                    .ToString()
                    .ToLower()
                    .Contains("field name"))
                {
                    return i;
                }
            }

            return 1;
        }

        private static void CopyHeaderBetweenWorkbooks(Workbook newWB, Worksheet sourceWS)
        {
            Range headerRange = sourceWS.UsedRange.Offset[GetIdentifyingRow(sourceWS) - 1, 1].Resize[1, sourceWS.UsedRange.Columns.Count - 1];
            headerRange.Copy((Range)outputWorksheet.Cells.Item[1, 1]);
            newWB.Worksheets[1].Delete();
        }

        /// <summary>
        /// Searches for match between <paramref name="fieldName"/> and first row values.
        /// </summary>
        /// <param name="fieldName">Field name to search.</param>
        /// <returns>Index of matched column.</returns>
        private static int FindColumnIndex(string fieldName)
        {
            for (int i = 1; i <= outputWorksheet.UsedRange.Columns.Count; i++)
            {
                if (outputWorksheet.Cells[1, i].Value2 == fieldName)
                {
                    return i;
                }
            }

            return 1;
        }

        /// <summary>
        /// Gets first empty row in <see cref="Worksheet"/> used range.
        /// </summary>
        /// <returns>Index of row.</returns>
        private static int GetLastRowInWorksheet()
        {
            return outputWorksheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row + 1; // .Rows[outputWorksheet.UsedRange.Rows.Count].Row + 1;
        }
    }
}