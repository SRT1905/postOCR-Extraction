using Microsoft.Office.Interop.Word;
using System;

namespace SmartOCR
{
    internal class WordApplication
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
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
            GC.Collect();
        }

        public static Document OpenWordDocument(string file_path)
        {
            Document document = GetWordApplication().Documents.Open(file_path);
            document.Activate();
            return document;
        }

        public static void CloseActiveWordDocument()
        {
            Application application = GetWordApplication();
            application.ActiveDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
        }

        public static void CloseDocument(Document document)
        {
            document.Close(WdSaveOptions.wdDoNotSaveChanges);
        }
    }
}