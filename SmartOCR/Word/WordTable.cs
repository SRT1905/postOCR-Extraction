using Microsoft.Office.Interop.Word;
using System;

namespace SmartOCR
{
    class WordTable
    {
        private readonly string[][] _cells;
        public string this[int Row, int Column]
        {
            get
            {
                if (_cells.Length > Row || Row < 0)
                {
                    return null;
                }
                if (_cells[Row].Length > Column || Column < 0)
                {
                    return null;
                }
                return _cells[Row][Column];
            }
        }
        public WordTable(Table wordTable)
        {
            _cells = new string[wordTable.Rows.Count][];
            for (int i = 0; i < wordTable.Rows.Count; i++)
            {
                _cells[i] = new string[wordTable.Columns.Count];
            }
            foreach (Cell cell in wordTable.Range.Cells)
            {
                _cells[cell.RowIndex - 1][cell.ColumnIndex - 1] = RemoveInvalidChars(cell.Range.Text);
            }
        }
        private string RemoveInvalidChars(string check_string)
        {
            string[] separators = new string[] { "\r", "\a", "\t", "\f" };
            string[] temp = check_string.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return string.Join("", temp).Replace("\v", " ");
        }
    }
}
