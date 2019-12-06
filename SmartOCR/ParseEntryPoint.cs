using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SmartOCR
{
    internal class ParseEntryPoint : IDisposable
    {
        private Dictionary<string, object> config_data = new Dictionary<string, object>();
        private string doc_type;
        private string output_location = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        private Workbook output_wb;

        private HashSet<string> valid_doc_types = Utilities.valid_document_types;

        private List<string> valid_files;

        public ParseEntryPoint(string type, IEnumerable<string> files)
        {
            doc_type = type;
            valid_files = GetValidFiles(files);
            config_data = new ExcelConfigParser().ParseConfig(doc_type);
            output_wb = ExcelOutputWorkbook.GetOutputWorkbook(doc_type);
        }

        public void Dispose()
        {
            config_data = null;
            doc_type = null;
            output_location = null;
            output_wb = null;
            valid_doc_types = null;
            valid_files = null;
            WordApplication.ExitWordApplication();
            ExcelApplication.ExitExcelApplication();
        }

        public bool TryGetData()
        {
            if (valid_files.Count == 0 || !valid_doc_types.Contains(doc_type))
            {
                return false;
            }
            GetDataFromFiles();
            return true;
        }

        private void GetDataFromFiles()
        {
            foreach (var item in valid_files)
            {
                Dictionary<string, string> result = GetResultFromFile(item);
                ExcelOutputWorkbook.ReturnValuesToWorksheet(result);
            }
            output_wb.SaveAs(output_location);
        }

        private Dictionary<string, string> GetResultFromFile(string item)
        {
            Document document = WordApplication.OpenWordDocument(item);
            WordReader reader = new WordReader(document);
            WordParser wordParser = new WordParser(reader.line_mapping);
            return wordParser.ParseDocument(config_data);
        }

        private List<string> GetValidFiles(IEnumerable<string> files)
        {
            return (from string file_path in files
                    where File.Exists(file_path)
                    select file_path).ToList();
        }
    }
}