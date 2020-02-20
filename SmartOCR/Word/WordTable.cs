using Microsoft.Office.Interop.Word;
using System;
using System.Collections;

namespace SmartOCR
{
    public class WordTable
    {
        private readonly string[][] _cells;
        public string this[int Row, int Column]
        {
            get
            {
                Row = Row < 0
                    ? RowCount + Row
                    : Row;
                Column = Column < 0
                    ? ColumnCount + Column
                    : Column;
                if (!(0 <= Row && Row <= _cells.Length ))
                {
                    return null;
                }
                if (!(0 <= Column && Column <= _cells[Row].Length))
                {
                    return null;
                }
                return _cells[Row][Column];
            }
        }
        public int RowCount
        {
            get
            {
                return _cells.Length;
            }
        }

        public int ColumnCount
        {
            get
            {
                return _cells[0].Length;
            }
        }

        public WordTable(Table wordTable)
        {
            if (wordTable == null)
            {
                throw new ArgumentNullException(nameof(wordTable));
            }
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
