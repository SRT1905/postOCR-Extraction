using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace SmartOCR
{
    /// <summary>
    /// Represents an entry point for processing input from command prompt or start form.
    /// </summary>
    public sealed class ParseEntryPoint : IDisposable
    {
        #region Constants
        /// <summary>
        /// Name of output file with extension.
        /// </summary>
        private const string outputFileName = "output.xlsx";
        #endregion

        #region Fields
        /// <summary>
        /// Object that describes config fields and their search expressions.
        /// </summary>
        private ConfigData configData = new ConfigData();
        /// <summary>
        /// Path to output file (default location equals assembly location).
        /// </summary>
        private string outputLocation = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), outputFileName);
        /// <summary>
        /// Instance of output Excel workbook.
        /// </summary>
        private Workbook outputWB;
        /// <summary>
        /// Collection of files to process.
        /// </summary>
        private List<string> validFiles = new List<string>();
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of ParseEntryPoint that uses internal config data and default output location.
        /// </summary>
        /// <param name="files">Collection of files to process.</param>
        public ParseEntryPoint(List<string> files)
        {
            if (files == null)
            {
                throw new ArgumentNullException(nameof(files));
            }
            validFiles = GetValidFiles(files);
            configData = new ConfigParser().ParseConfig();
            outputWB = ExcelOutputWorkbook.GetOutputWorkbook();
        }
        /// <summary>
        /// Initializes a new instance of ParseEntryPoint that uses external config data and default output location.
        /// </summary>
        /// <param name="files">Collection of files to process.</param>
        /// <param name="configFile">Path to external Excel workbook with config data.</param>
        public ParseEntryPoint(List<string> files, string configFile)
        {
            if (files == null)
            {
                throw new ArgumentNullException(nameof(files));
            }
            validFiles = GetValidFiles(files);
            configData = new ConfigParser(configFile).ParseConfig();
            outputWB = ExcelOutputWorkbook.GetOutputWorkbook();
        }
        /// <summary>
        /// Initializes a new instance of ParseEntryPoint that uses external config data and default output location.
        /// </summary>
        /// <param name="files">Collection of files to process.</param>
        /// <param name="configFile">Path to external Excel workbook with config data.</param>
        /// <param name="outputFile">Path to existing output Excel workbook.</param>
        public ParseEntryPoint(List<string> files, string configFile, string outputFile) : this(files, configFile)
        {
            outputLocation = outputFile;
        }
        #endregion

        #region Public methods
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
        /// Tries to get data, described in config data, from provided files.
        /// </summary>
        /// <returns>Indicator that processing was successful.</returns>
        public bool TryGetData()
        {
            if (validFiles.Count == 0)
            {
                return false;
            }
            GetDataFromFiles();
            return true;
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Sets instance fields values to null.
        /// </summary>
        private void DisposeFields()
        {
            configData = null;
            outputLocation = null;
            outputWB = null;
            validFiles = null;
        }
        /// <summary>
        /// Gets data from each provided file and returns result to output Excel workbook.
        /// </summary>
        private void GetDataFromFiles()
        {
            for (int i = 0; i < validFiles.Count; i++)
            {
                Dictionary<string, string> result = GetResultFromFile(validFiles[i]);
                ExcelOutputWorkbook.ReturnValuesToWorksheet(result);
            }
            outputWB.SaveAs(outputLocation);
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
            var wordParser = new WordParser(reader, configData);
            reader.Dispose();
            return wordParser.ParseDocument();
        }
        /// <summary>
        /// Checks whether provided file paths are valid - they exists and do not start with '~' symbol.
        /// </summary>
        /// <param name="filePaths">Collection of provided file paths.</param>
        /// <returns>Collection of valid files.</returns>
        private List<string> GetValidFiles(List<string> filePaths)
        {
            var files = new List<string>();
            for (int i = 0; i < filePaths.Count; i++)
            {
                string singleFile = filePaths[i];
                if (File.Exists(singleFile) && !singleFile.Contains("~"))
                {
                    files.Add(singleFile);
                }
            }
            return files;
        }
        #endregion
    }
}