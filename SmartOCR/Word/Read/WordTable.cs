namespace SmartOCR
{
    using System;
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Performs as a container of Word table data.
    /// </summary>
    public class WordTable
    {
        private readonly string[][] cells;
        private readonly ParagraphContainer anchor;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTable"/> class.
        /// </summary>
        /// <param name="wordTable">A reference to Word table.</param>
        public WordTable(Table wordTable)
        {
            if (wordTable == null)
            {
                throw new ArgumentNullException(nameof(wordTable));
            }

            this.cells = InitializeCells(wordTable);
            this.anchor = this.GetFirstCellDescription(wordTable);
        }

        /// <summary>
        /// Gets paragraph description of top-left table cell.
        /// </summary>
        public ParagraphContainer Anchor
        {
            get
            {
                return this.anchor;
            }
        }

        /// <summary>
        /// Gets count of rows in table.
        /// </summary>
        public int RowCount
        {
            get
            {
                return this.cells.Length;
            }
        }

        /// <summary>
        /// Gets count of columns in table.
        /// </summary>
        public int ColumnCount
        {
            get
            {
                return this.cells[0].Length;
            }
        }

        /// <summary>
        /// Returns a string value that represents a cell in a table.
        /// </summary>
        /// <param name="row">The number of the row in the table to return. May be negative - then row is taken from table end.</param>
        /// <param name="column">The number of the cell in the table to return. May be negative - then cell is taken from row end.</param>
        /// <returns>A cell value.</returns>
        public string this[int row, int column]
        {
            get
            {
                return this.ValidateIndexers(ref row, ref column)
                    ? this.cells[row][column]
                    : null;
            }
        }

        private static string[][] InitializeCells(Table wordTable)
        {
            string[][] tableCells = InitializeRows(wordTable);
            InitializeColumns(wordTable, tableCells);
            return tableCells;
        }

        private static void InitializeColumns(Table wordTable, string[][] tableCells)
        {
            foreach (Cell cell in wordTable.Range.Cells)
            {
                tableCells[cell.RowIndex - 1][cell.ColumnIndex - 1] = RemoveInvalidChars(cell.Range.Text);
            }
        }

        private static string[][] InitializeRows(Table wordTable)
        {
            string[][] tableCells = new string[wordTable.Rows.Count][];
            for (int i = 0; i < wordTable.Rows.Count; i++)
            {
                tableCells[i] = new string[wordTable.Columns.Count];
            }

            return tableCells;
        }

        private static string RemoveInvalidChars(string checkString)
        {
            string[] separators = new string[] { "\r", "\a", "\t", "\f" };
            string[] temp = checkString.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(string.Empty, temp).Replace("\v", " ");
        }

        private static int ValidateSingleIndexer(int indexer, int marginValue)
        {
            indexer = indexer < 0
                ? marginValue + indexer
                : indexer;
            return indexer;
        }

        private ParagraphContainer GetFirstCellDescription(Table wordTable)
        {
            for (int i = 0; i < this.RowCount; i++)
            {
                for (int j = 0; j < this.ColumnCount; j++)
                {
                    if (!string.IsNullOrEmpty(this[i, j]))
                    {
                        return new ParagraphContainer(wordTable.Cell(i + 1, j + 1).Range);
                    }
                }
            }

            return null;
        }

        private bool ValidateIndexers(ref int row, ref int column)
        {
            row = ValidateSingleIndexer(row, this.RowCount);
            column = ValidateSingleIndexer(column, this.ColumnCount);
            return row >= 0 && row <= this.cells.Length && column >= 0 && column <= this.cells[row].Length;
        }
    }
}
