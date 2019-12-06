using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace SmartOCR
{
    class ExcelApplication
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            GC.Collect();
        }

        public static Workbook OpenExcelWorkbook(string path)
        {
            Application app = GetExcelApplication();
            foreach (Workbook item in app.Workbooks)
            {
                if (item.Path == path)
                    return item;
            }
            Workbook workbook = app.Workbooks.Open(path);
            workbook.Activate();
            return workbook;
        }

        public static void CloseActiveExcelWorkbook()
        {
            Application app = GetExcelApplication();
            string path = app.ActiveWorkbook.Path;
            app.ActiveWorkbook.Close(false);
            if (Path.GetDirectoryName(path) == Path.GetTempPath())
                File.Delete(path);
        }
    }
}
