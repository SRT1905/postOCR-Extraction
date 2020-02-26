namespace SmartOCR
{
    using System.Globalization;
    using System.IO;
    using Microsoft.Office.Interop.Excel;

    public class ConfigParser // TODO: add summary.
    {
        public ConfigParser()
        {
            if (ConfigWorkbook == null)
            {
                ConfigWorkbook = this.GetInternalConfigWorkbook();
            }
        }

        public ConfigParser(string configFile)
        {
            ConfigWorkbook = this.GetExternalConfigWorkbook(configFile);
        }

        public static Workbook ConfigWorkbook { get; private set; }

        public ConfigData ParseConfig()
        {
            return this.GetConfigData(ConfigWorkbook.Worksheets[1]);
        }

        private static ConfigField GetFieldDefinition(Worksheet sourceWS, long headerRow, long fieldColumn)
        {
            string fieldName = sourceWS.Cells.Item[headerRow, fieldColumn].Value2;
            if (string.IsNullOrEmpty(fieldName))
            {
                return null;
            }

            string valueType = sourceWS.Cells.Item[headerRow + 1, fieldColumn].Value2;
            var field = new ConfigField(fieldName, valueType);
            field.ParseFieldExpression(sourceWS.Cells.Item[headerRow + 2, fieldColumn].Value2);
            long lastRow = sourceWS.Cells.Item[sourceWS.Rows.Count, fieldColumn].End[XlDirection.xlUp].Row;
            for (long i = headerRow + 3; i <= lastRow; i++)
            {
                field.AddSearchExpression(new ConfigExpression(valueType, sourceWS.Cells.Item[i, fieldColumn].Value2));
            }

            return field;
        }

        private Workbook GetExternalConfigWorkbook(string path)
        {
            return ExcelApplication.OpenExcelWorkbook(path);
        }

        private Workbook GetInternalConfigWorkbook()
        {
            string tempPath = Path.GetTempFileName();
            File.WriteAllBytes(tempPath, ConfigContainer.config);
            return ExcelApplication.OpenExcelWorkbook(tempPath);
        }

        private ConfigData GetConfigData(Worksheet sourceWS)
        {
            ConfigData data = new ConfigData();
            for (long headerRow = 1; headerRow <= sourceWS.UsedRange.Columns[1].Rows.Count; headerRow++)
            {
                if (sourceWS.Cells.Item[headerRow, 1].Value2.ToLower(CultureInfo.CurrentCulture).Contains("field name"))
                {
                    for (int fieldIndex = 2; fieldIndex <= sourceWS.UsedRange.Rows[headerRow].Columns.Count; fieldIndex++)
                    {
                        data.AddField(GetFieldDefinition(sourceWS, headerRow, fieldIndex));
                    }

                    return data;
                }
            }

            return data;
        }
    }
}