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

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigParser"/> class.
        /// Sets externally stored Excel workbook as source of config data.
        /// </summary>
        /// <param name="configFile">Path to config Excel workbook.</param>
        public ConfigParser(string configFile)
        {
            Utilities.Debug("Opening configuration file.");
            configWorkbook = this.GetExternalConfigWorkbook(configFile);
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
            SetSimilarityAlgorithm(configWorkbook.Worksheets[1]);
            return this.GetConfigData(configWorkbook.Worksheets[1]);
        }

        private static void SetSimilarityAlgorithm(Worksheet sourceWs)
        {
            for (var row = 1; row <= sourceWs.UsedRange.Rows.Count; row++)
            {
                string algorithmName = sourceWs.UsedRange.Cells.Item[row, 1].Value2;
                if (!algorithmName.ToLower().StartsWith("similarity algorithm"))
                {
                    continue;
                }

                ISimilarityAlgorithm algorithm =
                    SimilarityAlgorithmSelector.GetAlgorithm(sourceWs.UsedRange.Cells.Item[row, 2].Value2);
                if (algorithm == null)
                {
                    continue;
                }

                Utilities.Debug($"String similarity algorithm is {sourceWs.UsedRange.Cells.Item[row, 2].Value2}.");
                SimilarityDescription.SimilarityAlgorithm = algorithm;
                return;
            }

            Utilities.Debug("Default string similarity algorithm is used.", 2);
        }

        /// <summary>
        /// Gets definition of single config field and its search expressions.
        /// </summary>
        /// <param name="sourceWs">An Excel worksheet with config data.</param>
        /// <param name="headerRow">Index of row on worksheet with field names.</param>
        /// <param name="fieldColumn">Index of column where field definition is contained.</param>
        /// <returns>An instance of <see cref="ConfigField"/> that describes search field and its expressions.</returns>
        private static ConfigField GetFieldDefinition(Worksheet sourceWs, int headerRow, int fieldColumn)
        {
            string fieldName = sourceWs.Cells.Item[headerRow, fieldColumn].Value2;
            if (string.IsNullOrEmpty(fieldName))
            {
                return null;
            }

            Utilities.Debug($"Found field '{fieldName}'.", 2);
            return InitializeAndPopulateConfigField(sourceWs, headerRow, fieldColumn, fieldName, sourceWs.Cells.Item[GetValueTypeRow(sourceWs, headerRow), fieldColumn].Value2);
        }

        private static ConfigField InitializeAndPopulateConfigField(Worksheet sourceWs, int headerRow, int fieldColumn, string fieldName, string valueType)
        {
            ConfigField field = InitializeConfigField(fieldName, valueType, sourceWs.Cells.Item[GetFieldExpressionRow(sourceWs, headerRow), fieldColumn].Value2);
            SetGridCoordinates(field, sourceWs.Cells.Item[GetGridCoordinatesRow(sourceWs, headerRow), fieldColumn].Value2);
            AddSearchExpressionsToField(field, sourceWs, headerRow, fieldColumn);
            PrintDebugMessage(field);
            return field;
        }

        private static void PrintDebugMessage(ConfigField field)
        {
            string message = $"Value type: {field.ValueType}, {(field.UseSoundex ? "Soundex" : "Regular")} expression: {field.TextExpression}, Count of search expressions: {field.Expressions.Count}";
            Utilities.Debug(message, 3);
        }

        private static void AddSearchExpressionsToField(ConfigField field, Worksheet sourceWs, int headerRow, int fieldColumn)
        {
            for (int row = GetSearchValuesRow(sourceWs, headerRow); row <= sourceWs.Cells.Item[sourceWs.Rows.Count, fieldColumn].End[XlDirection.xlUp].Row; row++)
            {
                field.AddSearchExpression(new ConfigExpression(field.ValueType, sourceWs.Cells.Item[row, fieldColumn].Value2));
            }
        }

        private static void SetGridCoordinates(ConfigField field, string coordinatesValue)
        {
            field.ParseGridCoordinates(coordinatesValue);
        }

        private static int GetGridCoordinatesRow(Worksheet sourceWs, int headerRow)
        {
            return GetNonFieldNameIdentifierByTitle("grid coordinates", sourceWs, headerRow);
        }

        private static int GetSearchValuesRow(Worksheet sourceWs, int headerRow)
        {
            return GetNonFieldNameIdentifierByTitle("search values", sourceWs, headerRow);
        }

        private static int GetFieldExpressionRow(Worksheet sourceWs, int headerRow)
        {
            return GetNonFieldNameIdentifierByTitle("field expression", sourceWs, headerRow);
        }

        private static int GetValueTypeRow(Worksheet sourceWs, int headerRow)
        {
            return GetNonFieldNameIdentifierByTitle("value type", sourceWs, headerRow);
        }

        private static int GetNonFieldNameIdentifierByTitle(string title, Worksheet sourceWs, int headerRow)
        {
            for (int i = headerRow; i <= sourceWs.Cells.Item[sourceWs.Rows.Count, 1].End[XlDirection.xlUp].Row; i++)
            {
                if (sourceWs.Cells.Item[i, 1].Value2.ToLower().StartsWith(title))
                {
                    return i;
                }
            }

            return 1;
        }

        private static ConfigField InitializeConfigField(string fieldName, string valueType, string fieldDescription)
        {
            var field = new ConfigField(fieldName, valueType);
            field.ParseFieldExpression(fieldDescription);
            return field;
        }

        private static ConfigData AddConfigFields(ConfigData data, Worksheet sourceWs, int headerRow)
        {
            for (int fieldIndex = 2; fieldIndex <= sourceWs.UsedRange.Rows[headerRow].Columns.Count; fieldIndex++)
            {
                data.AddField(GetFieldDefinition(sourceWs, headerRow, fieldIndex));
            }

            return data;
        }

        private static Workbook GetInternalConfigWorkbook()
        {
            string tempPath = Path.GetTempFileName();
            File.WriteAllBytes(tempPath, Resources.ConfigContainer.config);
            configWorkbook = ExcelApplication.OpenExcelWorkbook(tempPath);
            return configWorkbook;
        }

        private static bool DoesCellHasIdentifier(Worksheet sourceWs, int headerRow)
        {
            return (bool)sourceWs.Cells
                                 .Item[headerRow, 1]
                                 .Value2
                                 .ToLower()
                                 .Contains("field name");
        }

        private Workbook GetExternalConfigWorkbook(string path)
        {
            return ExcelApplication.OpenExcelWorkbook(path);
        }

        private ConfigData GetConfigData(Worksheet sourceWs)
        {
            Utilities.Debug("Getting config data from first worksheet.");
            var data = new ConfigData();
            for (int headerRow = 1; headerRow <= sourceWs.UsedRange.Columns[1].Rows.Count; headerRow++)
            {
                if (!DoesCellHasIdentifier(sourceWs, headerRow))
                {
                    continue;
                }

                Utilities.Debug($"Found 'Field name' identifier at row {headerRow}.", 1);
                return AddConfigFields(data, sourceWs, headerRow);
            }

            return data;
        }
    }
}