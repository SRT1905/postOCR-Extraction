using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace SmartOCR
{
    internal class ExcelApplication
    {
        private static Application instance;

        public static Application GetExcelApplication()
        {
            if (instance == null)
            {
                instance = new Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false
                };
            }
            return instance;
        }

        public static void ExitExcelApplication()
        {
            Application app = GetExcelApplication();
            app.Quit();
            if (app != null)
            {
                Marshal.ReleaseComObject(app);
            }
            GC.Collect();
        }

        public static Workbook OpenExcelWorkbook(string path)
        {
            Application app = GetExcelApplication();
            return TryGetWorkbook(app.Workbooks, path);
        }

        public static void CloseActiveExcelWorkbook()
        {
            Application app = GetExcelApplication();
            app.ActiveWorkbook.Close(XlSaveAction.xlDoNotSaveChanges);
            TryDeleteTempFile(app.ActiveWorkbook.Path);
        }

        private static void TryDeleteTempFile(string path)
        {
            if (Path.GetDirectoryName(path) == Path.GetTempPath())
            {
                File.Delete(path);
            }
        }

        private static Workbook TryGetWorkbook(Workbooks workbooks, string path)
        {
            try
            {
                return workbooks[path];
            }
            catch (Exception)
            {
                return workbooks.Open(path);
            }
        }
    }
}