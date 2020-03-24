namespace SmartOCR.Config
{
    using System.IO;
    using Microsoft.Office.Interop.Excel;
    using SmartOCR.Excel;
    using SmartOCR.Search;
    using SmartOCR.Search.SimilarityAlgorithms;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

    /// <summary>
    /// Performs collection and parsing of config data from config Excel workbook.
    /// </summary>
    public class ConfigParser
    {
        private static Workbook configWorkbook;
        private readonly Worksheet worksheet;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigParser"/> class.
        /// Sets externally stored Excel workbook as source of config data.
        /// </summary>
        /// <param name="configFile">Path to config Excel workbook.</param>
        public ConfigParser(string configFile)
        {
            Utilities.Debug("Opening configuration file.");
            configWorkbook = this.GetExternalConfigWorkbook(configFile);
            this.worksheet = configWorkbook.Worksheets[1];
        }

        /// <summary>
        /// Gets single reference to Excel workbook with config data.
        /// </summary>
        public static Workbook ConfigWorkbook => configWorkbook ?? GetInternalConfigWorkbook();

        /// <summary>
        /// Gets config data from first worksheet on <see cref="ConfigWorkbook"/>.
        /// </summary>
        /// <returns>An instance of <see cref="ConfigData"/>.</returns>
        public ConfigData ParseConfig()
        {
            this.SetSimilarityAlgorithm();
            return this.GetConfigData();
        }

        private static void PrintDebugMessage(ConfigField field)
        {
            string message = $"Value type: {field.ValueType}, {(field.UseSoundex ? "Soundex" : "Regular")} expression: {field.TextExpression}, Count of search expressions: {field.Expressions.Count}";
            Utilities.Debug(message, 3);
        }

        private static void SetGridCoordinates(ConfigField field, string coordinatesValue)
        {
            field.ParseGridCoordinates(coordinatesValue);
        }

        private static ConfigField InitializeConfigField(string fieldName, string valueType, string fieldDescription)
        {
            var field = new ConfigField(fieldName, valueType);
            field.ParseFieldExpression(fieldDescription);
            return field;
        }

        private static Workbook GetInternalConfigWorkbook()
        {
            string tempPath = Path.GetTempFileName();
            File.WriteAllBytes(tempPath, Resources.ConfigContainer.config);
            configWorkbook = ExcelApplication.OpenExcelWorkbook(tempPath);
            return configWorkbook;
        }

        /// <summary>
        /// Gets definition of single config field and its search expressions.
        /// </summary>
        /// <param name="headerRow">Index of row on worksheet with field names.</param>
        /// <param name="fieldColumn">Index of column where field definition is contained.</param>
        /// <returns>An instance of <see cref="ConfigField"/> that describes search field and its expressions.</returns>
        private ConfigField GetFieldDefinition(int headerRow, int fieldColumn)
        {
            string fieldName = this.worksheet.Cells.Item[headerRow, fieldColumn].Value2;
            if (string.IsNullOrEmpty(fieldName))
            {
                return null;
            }

            Utilities.Debug($"Found field '{fieldName}'.", 2);
            return this.InitializeAndPopulateConfigField(
                headerRow,
                fieldColumn,
                fieldName,
                this.worksheet.Cells.Item[this.GetValueTypeRow(headerRow), fieldColumn].Value2);
        }

        private ConfigField InitializeAndPopulateConfigField(int headerRow, int fieldColumn, string fieldName, string valueType)
        {
            ConfigField field = InitializeConfigField(fieldName, valueType, this.worksheet.Cells.Item[this.GetFieldExpressionRow(headerRow), fieldColumn].Value2);
            SetGridCoordinates(field, this.worksheet.Cells.Item[this.GetGridCoordinatesRow(headerRow), fieldColumn].Value2);
            this.AddSearchExpressionsToField(field, headerRow, fieldColumn);
            PrintDebugMessage(field);
            return field;
        }

        private void AddSearchExpressionsToField(ConfigField field, int headerRow, int fieldColumn)
        {
            for (int row = this.GetSearchValuesRow(headerRow); row <= this.worksheet.Cells.Item[this.worksheet.Rows.Count, fieldColumn].End[XlDirection.xlUp].Row; row++)
            {
                field.AddSearchExpression(new ConfigExpression(field.ValueType, this.worksheet.Cells.Item[row, fieldColumn].Value2));
            }
        }

        private int GetGridCoordinatesRow(int headerRow)
        {
            return this.GetNonFieldNameIdentifierByTitle("grid coordinates", headerRow);
        }

        private int GetSearchValuesRow(int headerRow)
        {
            return this.GetNonFieldNameIdentifierByTitle("search values", headerRow);
        }

        private int GetFieldExpressionRow(int headerRow)
        {
            return this.GetNonFieldNameIdentifierByTitle("field expression", headerRow);
        }

        private int GetValueTypeRow(int headerRow)
        {
            return this.GetNonFieldNameIdentifierByTitle("value type", headerRow);
        }

        private int GetNonFieldNameIdentifierByTitle(string title, int headerRow)
        {
            for (int i = headerRow; i <= this.worksheet.Cells.Item[this.worksheet.Rows.Count, 1].End[XlDirection.xlUp].Row; i++)
            {
                if (this.worksheet.Cells.Item[i, 1].Value2.ToLower().StartsWith(title))
                {
                    return i;
                }
            }

            return 1;
        }

        private void SetSimilarityAlgorithm()
        {
            for (var row = 1; row <= this.worksheet.UsedRange.Rows.Count; row++)
            {
                string algorithmName = this.worksheet.UsedRange.Cells.Item[row, 1].Value2;
                if (!algorithmName.ToLower().StartsWith("similarity algorithm"))
                {
                    continue;
                }

                ISimilarityAlgorithm algorithm =
                    SimilarityAlgorithmSelector.GetAlgorithm(this.worksheet.UsedRange.Cells.Item[row, 2].Value2);
                if (algorithm == null)
                {
                    continue;
                }

                Utilities.Debug($"String similarity algorithm is {this.worksheet.UsedRange.Cells.Item[row, 2].Value2}.");
                SimilarityDescription.SimilarityAlgorithm = algorithm;
                return;
            }

            Utilities.Debug("Default string similarity algorithm is used.", 2);
        }

        private ConfigData AddConfigFields(ConfigData data, int headerRow)
        {
            for (int fieldIndex = 2; fieldIndex <= this.worksheet.UsedRange.Rows[headerRow].Columns.Count; fieldIndex++)
            {
                data.AddField(this.GetFieldDefinition(headerRow, fieldIndex));
            }

            return data;
        }

        private bool DoesCellHasIdentifier(int headerRow)
        {
            return (bool)this.worksheet.Cells
                                       .Item[headerRow, 1]
                                       .Value2
                                       .ToLower()
                                       .Contains("field name");
        }

        private Workbook GetExternalConfigWorkbook(string path)
        {
            return ExcelApplication.OpenExcelWorkbook(path);
        }

        private ConfigData GetConfigData()
        {
            Utilities.Debug("Getting config data from first worksheet.");
            var data = new ConfigData();
            for (int headerRow = 1; headerRow <= this.worksheet.UsedRange.Columns[1].Rows.Count; headerRow++)
            {
                if (!this.DoesCellHasIdentifier(headerRow))
                {
                    continue;
                }

                Utilities.Debug($"Found 'Field name' identifier at row {headerRow}.", 1);
                return this.AddConfigFields(data, headerRow);
            }

            return data;
        }
    }
}