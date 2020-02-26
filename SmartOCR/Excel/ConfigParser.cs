using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.IO;

namespace SmartOCR
{

    public class ConfigParser // TODO: add summary.
    {
        #region Static properties
        public static Workbook ConfigWorkbook { get; private set; }
        #endregion

        #region Constructors
        public ConfigParser()
        {
            if (ConfigWorkbook == null)
            {
                ConfigWorkbook = GetInternalConfigWorkbook();
            }
        }
        public ConfigParser(string configFile)
        {
            ConfigWorkbook = GetExternalConfigWorkbook(configFile);
        }
        #endregion

        #region Public methods
        public ConfigData ParseConfig()
        {
            return GetConfigData(ConfigWorkbook.Worksheets[1]);
        }
        #endregion

        #region Private methods
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
            var data = new ConfigData();
            long headerRow;
            for (headerRow = 1; headerRow <= sourceWS.Cells.Item[sourceWS.Rows.Count, 1].End[XlDirection.xlUp].Row; headerRow++)
            {
                string cellValue = sourceWS.Cells.Item[headerRow, 1].Value2;
                if (cellValue.ToLower(CultureInfo.CurrentCulture).Contains("field name"))
                {
                    break;
                }
            }

            long lastColumn = sourceWS.Cells.Item[headerRow, sourceWS.Columns.Count].End[XlDirection.xlToLeft].Column;
            for (int i = 2; i <= lastColumn; i++)
            {
                ConfigField field = GetFieldDefinition(sourceWS, headerRow, i);
                data.AddField(field);
            }
            return data;
        }
        #endregion

        #region Private static methods        
        private static ConfigField GetFieldDefinition(Worksheet sourceWS, long headerRow, long fieldColumn)
        {
            string fieldName = sourceWS.Cells.Item[headerRow, fieldColumn].Value2;
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
        #endregion
    }
}