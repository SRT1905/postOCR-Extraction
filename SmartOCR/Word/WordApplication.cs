using Microsoft.Office.Interop.Word;
using System;

namespace SmartOCR
{
    class WordApplication
    {
        private static Application instance;

        public static Application GetWordApplication()
        {
            if (instance == null)
            {
                instance = new Application
                {
                    Visible = false,
                    DisplayAlerts = WdAlertLevel.wdAlertsNone,
                    ScreenUpdating = false
                };
            }
            return instance;
        }

        public static void ExitWordApplication()
        {
            Application app = GetWordApplication();
            app.Quit(WdSaveOptions.wdDoNotSaveChanges);
            if (app != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            GC.Collect();
        }

        public static Document OpenWordDocument(string file_path)
        {
            Application application = GetWordApplication();
            Document document = application.Documents.Open(file_path);
            document.Activate();
            return document;
        }

        public static void CloseActiveWordDocument()
        {
            Application application = GetWordApplication();
            application.ActiveDocument.Close(false);
        }
    }
}
