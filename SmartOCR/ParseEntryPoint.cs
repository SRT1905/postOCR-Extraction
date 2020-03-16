namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Represents an entry point for processing input from command prompt or start form.
    /// </summary>
    public sealed class ParseEntryPoint : IDisposable
    {
        /// <summary>
        /// Name of output file with extension.
        /// </summary>
        private const string OutputFileName = "output.xlsx";

        /// <summary>
        /// Object that describes config fields and their search expressions.
        /// </summary>
        private ConfigData configData = new ConfigData();

        /// <summary>
        /// Path to output file (default location equals assembly location).
        /// </summary>
        private string outputLocation = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), OutputFileName);

        /// <summary>
        /// Instance of output Excel workbook.
        /// </summary>
        private Workbook outputWB;

        /// <summary>
        /// Collection of files to process.
        /// </summary>
        private List<string> validFiles = new List<string>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ParseEntryPoint"/> class.
        /// Instance uses external config data and default output location.
        /// </summary>
        /// <param name="files">Collection of files to process.</param>
        /// <param name="configFile">Path to external Excel workbook with config data.</param>
        public ParseEntryPoint(List<string> files, string configFile)
        {
            if (files == null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            this.InitializeFields(files, configFile);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ParseEntryPoint"/> class.
        /// Instance uses external config data and default output location.
        /// </summary>
        /// <param name="files">Collection of files to process.</param>
        /// <param name="configFile">Path to external Excel workbook with config data.</param>
        /// <param name="outputFile">Path to existing output Excel workbook.</param>
        public ParseEntryPoint(List<string> files, string configFile, string outputFile)
            : this(files, configFile)
        {
            this.outputLocation = outputFile;
        }

        /// <summary>
        /// Disposes of field values and open Office applications.
        /// </summary>
        public void Dispose()
        {
            this.DisposeFields();
            WordApplication.ExitWordApplication();
            ExcelApplication.ExitExcelApplication();
        }

        /// <summary>
        /// Tries to get data, described in config data, from provided files.
        /// </summary>
        /// <returns>Indicator that processing was successful.</returns>
        public bool TryGetData()
        {
            if (this.validFiles.Count != 0)
            {
                this.GetDataFromFiles();
            }

            return this.validFiles.Count != 0;
        }

        private static WordReader OpenAndReadDocument(string item)
        {
            var reader = new WordReader(WordApplication.OpenWordDocument(item));
            reader.ReadDocument();
            return reader;
        }

        private static Func<string, bool> IsFilePathValid()
        {
            return singlePath => File.Exists(singlePath) && !singlePath.Contains("~");
        }

        private void InitializeFields(List<string> files, string configFile)
        {
            this.validFiles = this.GetValidFiles(files);
            this.configData = new ConfigParser(configFile).ParseConfig();
            this.outputWB = ExcelOutputWorkbook.GetOutputWorkbook();
        }

        /// <summary>
        /// Sets instance fields values to null.
        /// </summary>
        private void DisposeFields()
        {
            this.configData = null;
            this.outputLocation = null;
            this.outputWB = null;
            this.validFiles = null;
        }

        /// <summary>
        /// Gets data from each provided file and returns result to output Excel workbook.
        /// </summary>
        private void GetDataFromFiles()
        {
            for (int i = 0; i < this.validFiles.Count; i++)
            {
                this.GetDataFromSingleFile(i);
            }

            this.outputWB.SaveAs(this.outputLocation);
            Utilities.Debug($"Excel output workbook is saved at location '{this.outputLocation}'.");
        }

        private void GetDataFromSingleFile(int fileIndex)
        {
            Dictionary<string, string> result = this.ProcessSingleFile(this.validFiles[fileIndex]);
            ExcelOutputWorkbook.ReturnValuesToWorksheet(result);
            Utilities.Debug($"File '{this.validFiles[fileIndex]}' was processed.", 1);
        }

        /// <summary>
        /// Performs processing of single document.
        /// </summary>
        /// <param name="wordFilePath">Path to single file.</param>
        /// <returns>Mapping between config field and found value.</returns>
        private Dictionary<string, string> ProcessSingleFile(string wordFilePath)
        {
            Utilities.Debug($"Processing file: '{wordFilePath}'.");
            WordParser wordParser;
            using (WordReader reader = OpenAndReadDocument(wordFilePath))
            {
                wordParser = new WordParser(reader, this.configData);
            }

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
            files.AddRange(filePaths.Where(IsFilePathValid()));
            Utilities.Debug($"Total number of files to process: {files.Count}");
            return files;
        }
    }
}