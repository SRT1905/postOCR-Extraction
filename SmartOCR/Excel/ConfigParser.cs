namespace SmartOCR
{
    using System.Globalization;
    using System.IO;
    using Microsoft.Office.Interop.Excel;

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
            configWorkbook = this.GetExternalConfigWorkbook(configFile);
        }

        /// <summary>
        /// Gets single reference to Excel workbook with config data.
        /// </summary>
        public static Workbook ConfigWorkbook
        {
            get
            {
                return configWorkbook ?? GetInternalConfigWorkbook();
            }

            private set
            {
                configWorkbook = value;
            }
        }

        /// <summary>
        /// Gets config data from first worksheet on <see cref="ConfigWorkbook"/>.
        /// </summary>
        /// <returns>An instance of <see cref="ConfigData"/>.</returns>
        public ConfigData ParseConfig()
        {
            return this.GetConfigData(configWorkbook.Worksheets[1]);
        }

        /// <summary>
        /// Gets definition of single config field and its search expressions.
        /// </summary>
        /// <param name="sourceWS">An Excel worksheet with config data.</param>
        /// <param name="headerRow">Index of row on worksheet with field names.</param>
        /// <param name="fieldColumn">Index of column where field definition is contained.</param>
        /// <returns>An instance of <see cref="ConfigField"/> that describes search field and its expressions.</returns>
        private static ConfigField GetFieldDefinition(Worksheet sourceWS, long headerRow, long fieldColumn)
        {
            string fieldName = sourceWS.Cells.Item[headerRow, fieldColumn].Value2;
            if (string.IsNullOrEmpty(fieldName))
            {
                return null;
            }

            return InitializeAndPopulateConfigField(sourceWS, headerRow, fieldColumn, fieldName, sourceWS.Cells.Item[headerRow + 1, fieldColumn].Value2);
        }

        private static ConfigField InitializeAndPopulateConfigField(Worksheet sourceWS, long headerRow, long fieldColumn, string fieldName, string valueType)
        {
            ConfigField field = InitializeConfigField(sourceWS.Cells.Item[headerRow + 2, fieldColumn].Value2, fieldName, valueType);
            AddSearchExpressionsToField(sourceWS, headerRow, fieldColumn, valueType, field);

            return field;
        }

        private static void AddSearchExpressionsToField(Worksheet sourceWS, long headerRow, long fieldColumn, string valueType, ConfigField field)
        {
            for (long i = headerRow + 3; i <= sourceWS.Cells.Item[sourceWS.Rows.Count, fieldColumn].End[XlDirection.xlUp].Row; i++)
            {
                field.AddSearchExpression(new ConfigExpression(valueType, sourceWS.Cells.Item[i, fieldColumn].Value2));
            }
        }

        private static ConfigField InitializeConfigField(string fieldDescription, string fieldName, string valueType)
        {
            var field = new ConfigField(fieldName, valueType);
            field.ParseFieldExpression(fieldDescription);
            return field;
        }

        private static ConfigData AddConfigFields(Worksheet sourceWS, ConfigData data, long headerRow)
        {
            for (int fieldIndex = 2; fieldIndex <= sourceWS.UsedRange.Rows[headerRow].Columns.Count; fieldIndex++)
            {
                data.AddField(GetFieldDefinition(sourceWS, headerRow, fieldIndex));
            }

            return data;
        }

        private static Workbook GetInternalConfigWorkbook()
        {
            string tempPath = Path.GetTempFileName();
            File.WriteAllBytes(tempPath, ConfigContainer.config);
            configWorkbook = ExcelApplication.OpenExcelWorkbook(tempPath);
            return configWorkbook;
        }

        private Workbook GetExternalConfigWorkbook(string path)
        {
            return ExcelApplication.OpenExcelWorkbook(path);
        }

        private ConfigData GetConfigData(Worksheet sourceWS)
        {
            ConfigData data = new ConfigData();
            for (long headerRow = 1; headerRow <= sourceWS.UsedRange.Columns[1].Rows.Count; headerRow++)
            {
                if (sourceWS.Cells.Item[headerRow, 1].Value2.ToLower(CultureInfo.CurrentCulture).Contains("field name"))
                {
                    return AddConfigFields(sourceWS, data, headerRow);
                }
            }

            return data;
        }
    }
}