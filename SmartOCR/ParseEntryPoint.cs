using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SmartOCR
{
    internal class ParseEntryPoint : IDisposable
    {
        private Dictionary<string, object> config_data;
        private string doc_type;
        private string output_location;
        private Workbook output_wb;
        private HashSet<string> valid_doc_types;
        private List<string> valid_files;

        public ParseEntryPoint()
        {
            config_data = new Dictionary<string, object>();
            output_location = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            valid_doc_types = Utilities.valid_document_types;
            valid_files = new List<string>();
        }

        public ParseEntryPoint(string type, IEnumerable<string> files) : this()
        {
            doc_type = type;
            valid_files = GetValidFiles(files);
            config_data = new ExcelConfigParser().ParseConfig(doc_type);
            output_wb = ExcelOutputWorkbook.GetOutputWorkbook(doc_type);
        }

        public ParseEntryPoint(string type, IEnumerable<string> files, string config_file) : this()
        {
            doc_type = type;
            valid_files = GetValidFiles(files);
            config_data = new ExcelConfigParser(config_file).ParseConfig(doc_type);
            output_wb = ExcelOutputWorkbook.GetOutputWorkbook(doc_type);
        }

        public ParseEntryPoint(string type, IEnumerable<string> files, string config_file, string output_file) : this(type, files, config_file)
        {
            output_location = output_file;
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
            foreach (string item in valid_files)
            {
                Dictionary<string, string> result = GetResultFromFile(item);
                ExcelOutputWorkbook.ReturnValuesToWorksheet(result);
            }
            output_wb.SaveAs(output_location);
        }

        private Dictionary<string, string> GetResultFromFile(string item)
        {
            Document document = WordApplication.OpenWordDocument(item);
            using (var reader = new WordReader(document))
            {
                reader.ReadDocument();
                WordParser wordParser = new WordParser(reader.line_mapping, config_data);
                return wordParser.ParseDocument();
            }
        }

        private List<string> GetValidFiles(IEnumerable<string> files)
        {
            return (from string file_path in files
                    where File.Exists(file_path) && !file_path.Contains("~")
                    select file_path).ToList();
        }
    }
}