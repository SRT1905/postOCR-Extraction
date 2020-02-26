namespace SmartOCR
{
    using System;
    using Microsoft.Office.Interop.Word;

    public class WordTable
    {
        private readonly string[][] cells;

        public WordTable(Table wordTable)
        {
            if (wordTable == null)
            {
                throw new ArgumentNullException(nameof(wordTable));
            }

            this.cells = new string[wordTable.Rows.Count][];
            for (int i = 0; i < wordTable.Rows.Count; i++)
            {
                this.cells[i] = new string[wordTable.Columns.Count];
            }

            foreach (Cell cell in wordTable.Range.Cells)
            {
                this.cells[cell.RowIndex - 1][cell.ColumnIndex - 1] = this.RemoveInvalidChars(cell.Range.Text);
            }
        }

        public int RowCount
        {
            get
            {
                return this.cells.Length;
            }
        }

        public int ColumnCount
        {
            get
            {
                return this.cells[0].Length;
            }
        }

        public string this[int row, int column]
        {
            get
            {
                row = row < 0
                    ? this.RowCount + row
                    : row;
                column = column < 0
                    ? this.ColumnCount + column
                    : column;
                if (!(row >= 0 && row <= this.cells.Length))
                {
                    return null;
                }

                if (!(column >= 0 && column <= this.cells[row].Length))
                {
                    return null;
                }

                return this.cells[row][column];
            }
        }

        private string RemoveInvalidChars(string checkString)
        {
            string[] separators = new string[] { "\r", "\a", "\t", "\f" };
            string[] temp = checkString.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(string.Empty, temp).Replace("\v", " ");
        }
    }
}
