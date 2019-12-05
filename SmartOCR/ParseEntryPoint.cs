using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace SmartOCR
{
    class ParseEntryPoint
    {
        public Dictionary<string, object> Config_data { get; }
        private readonly List<string> valid_files;
        private readonly List<string> valid_doc_types = new List<string>()
        {
            "Invoice"
        };

        public ParseEntryPoint(string doc_type, IEnumerable<string> files)
        {
            valid_files = new List<string>();
            foreach (string file_path in files)
            {
                if (File.Exists(file_path))
                {
                    valid_files.Add(file_path);
                }
            }
            if (valid_files.Count == 0)
            {
                return;
            }

            if (valid_doc_types.Contains(doc_type))
            {
                ConfigParser parser = new ConfigParser();
                Config_data = parser.ParseConfig(doc_type);
                
                Workbook output_wb = OutputWorkbook.GetOutputWorkbook(doc_type);

                foreach (string item in valid_files)
                {
                    Document document = WordApplication.OpenWordDocument(item);
                    WordReader reader = new WordReader(document);
                    reader.Dispose();
                    WordParser wordParser = new WordParser(reader.line_mapping);
                    Dictionary<string, string> result = wordParser.ParseDocument(Config_data);

                    OutputWorkbook.ReturnValuesToWorksheet(result);

                    output_wb.SaveAs(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location));
                }
                WordApplication.ExitWordApplication();
                ExcelApplication.ExitExcelApplication();
            }
        }
    }
}
