using Microsoft.Office.Interop.Word;
using System;

namespace SmartOCR
{
    /// <summary>
    /// Performs communication with MS Word application.
    /// </summary>
    internal static class WordApplication
    {
        /// <summary>
        /// Single instance of Word application.
        /// </summary>
        private static Application instance;

        /// <summary>
        /// Gets existing Word application or initializes a new one.
        /// </summary>
        /// <returns>Word application.</returns>
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

        /// <summary>
        /// Closes Word application without saving any changes.
        /// </summary>
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

        /// <summary>
        /// Opens document in Word application and makes document active.
        /// </summary>
        /// <param name="file_path">Path to document file.</param>
        /// <returns>Document representation.</returns>
        public static Document OpenWordDocument(string file_path)
        {
            Document document = GetWordApplication().Documents.Open(file_path);
            document.Activate();
            return document;
        }

        /// <summary>
        /// Gets active document in Word application and closes it.
        /// </summary>
        public static void CloseActiveWordDocument()
        {
            Application application = GetWordApplication();
            application.ActiveDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
        }

        /// <summary>
        /// Closes provided Word document without saving any changes.
        /// </summary>
        /// <param name="document"></param>
        public static void CloseDocument(Document document)
        {
            document.Close(WdSaveOptions.wdDoNotSaveChanges);
        }
    }
}