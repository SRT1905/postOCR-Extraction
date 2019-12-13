using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SmartOCR
{
    /// <summary>
    /// Represents an entry point for processing input from command prompt or start form.
    /// </summary>
    internal class ParseEntryPoint : IDisposable
    {
        /// <summary>
        /// Name of output file with extension.
        /// </summary>
        private const string output_file_name = "output.xlsx";

        /// <summary>
        /// Object that describes config fields and their search expressions.
        /// </summary>
        private ConfigData config_data = new ConfigData();
        
        private string doc_type;
        
        /// <summary>
        /// Path to output file (default location equals assembly location).
        /// </summary>
        private string output_location = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), output_file_name);
        
        /// <summary>
        /// Instance of output Excel workbook.
        /// </summary>
        private Workbook output_wb;
        
        /// <summary>
        /// Collection of supported document types.
        /// </summary>
        private HashSet<string> valid_doc_types = Utilities.valid_document_types;
        
        /// <summary>
        /// Collection of files to process.
        /// </summary>
        private List<string> valid_files = new List<string>();
        
        /// <summary>
        /// Initializes a new instance of ParseEntryPoint that uses internal config data and default output location.
        /// </summary>
        /// <param name="type">Document type.</param>
        /// <param name="files">Collection of files to process.</param>
        public ParseEntryPoint(string type, List<string> files)
        {
            doc_type = type;
            valid_files = GetValidFiles(files);
            config_data = new ConfigParser().ParseConfig(doc_type);
            output_wb = ExcelOutputWorkbook.GetOutputWorkbook(doc_type);
        }

        /// <summary>
        /// Initializes a new instance of ParseEntryPoint that uses external config data and default output location.
        /// </summary>
        /// <param name="type">Document type.</param>
        /// <param name="files">Collection of files to process.</param>
        /// <param name="config_file">Path to external Excel workbook with config data.</param>
        public ParseEntryPoint(string type, List<string> files, string config_file)
        {
            doc_type = type;
            valid_files = GetValidFiles(files);
            config_data = new ConfigParser(config_file).ParseConfig(doc_type);
            output_wb = ExcelOutputWorkbook.GetOutputWorkbook(doc_type);
        }
      
        /// <summary>
        /// Initializes a new instance of ParseEntryPoint that uses external config data and default output location.
        /// </summary>
        /// <param name="type">Document type.</param>
        /// <param name="files">Collection of files to process.</param>
        /// <param name="config_file">Path to external Excel workbook with config data.</param>
        /// <param name="output_file">Path to existing output Excel workbook.</param>
        public ParseEntryPoint(string type, List<string> files, string config_file, string output_file) : this(type, files, config_file)
        {
            output_location = output_file;
        }
        

        /// <summary>
        /// Tries to get data, described in config data, from provided files.
        /// </summary>
        /// <returns>Indicator that processing was successful.</returns>
        public bool TryGetData()
        {
            if (valid_files.Count == 0 || !valid_doc_types.Contains(doc_type))
            {
                return false;
            }
            GetDataFromFiles();
            return true;
        }
    
        /// <summary>
        /// Disposes of field values and open Office applications.
        /// </summary>
        public void Dispose()
        {
            DisposeFields();
            WordApplication.ExitWordApplication();
            ExcelApplication.ExitExcelApplication();
        }
        
        /// <summary>
        /// Gets data from each provided file and returns result to output Excel workbook.
        /// </summary>
        private void GetDataFromFiles()
        {
            for (int i = 0; i < valid_files.Count; i++)
            {
                Dictionary<string, string> result = GetResultFromFile(valid_files[i]);
                ExcelOutputWorkbook.ReturnValuesToWorksheet(result);
            }
            output_wb.SaveAs(output_location);
        }

        /// <summary>
        /// Performs processing of single document.
        /// </summary>
        /// <param name="item">Path to single file.</param>
        /// <returns>Mapping between config field and found value.</returns>
        private Dictionary<string, string> GetResultFromFile(string item)
        {
            Document document = WordApplication.OpenWordDocument(item);
            var reader = new WordReader(document);
            reader.ReadDocument();
            WordParser wordParser = new WordParser(reader.LineMapping, config_data);
            reader.Dispose();
            return wordParser.ParseDocument();
        }

        /// <summary>
        /// Checks whether provided file paths are valid - they exists and do not start with '~' symbol.
        /// </summary>
        /// <param name="file_paths">Collection of provided file paths.</param>
        /// <returns>Collection of valid files.</returns>
        private List<string> GetValidFiles(List<string> file_paths)
        {
            var files = new List<string>();
            for (int i = 0; i < file_paths.Count; i++)
            {
                string single_file = file_paths[i];
                if (File.Exists(single_file) && !single_file.Contains("~"))
                {
                    files.Add(single_file);
                }
            }
            return files;
        }
      
        /// <summary>
        /// Sets instance fields values to null.
        /// </summary>
        private void DisposeFields()
        {
            config_data = null;
            doc_type = null;
            output_location = null;
            output_wb = null;
            valid_doc_types = null;
            valid_files = null;
        }
    }
}