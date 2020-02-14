using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace SmartOCR
{
    internal class WordTable
    {
        private readonly List<List<string>> table_contents = new List<List<string>>();
        public string this[int Row, int Column]
        {
            get
            {
                return table_contents[Row][Column];
            }
        }

        public WordTable(Word.Table table)
        {
            InitializeTable(table);
        }
        private void InitializeTable(Word.Table table)
        {
            for (int i = 1; i <= table.Rows.Count; i++)
            {
                var row_list = new List<string>();
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    row_list.Add(table.Cell(i, j).Range.Text);
                }
                table_contents.Add(row_list);
            }
        }
    }
}
