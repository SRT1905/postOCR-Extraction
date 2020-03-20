namespace SmartOCR.Word.Read
{
    using System;
    using Microsoft.Office.Interop.Word;
    using Runtime = System.Runtime;

    /// <summary>
    /// Performs communication with MS Word application.
    /// </summary>
    public static class WordApplication
    {
        /// <summary>
        /// Single instance of Word application.
        /// </summary>
        private static Application instance;

        /// <summary>
        /// Closes provided Word document without saving any changes.
        /// </summary>
        /// <param name="document">A Word document.</param>
        public static void CloseDocument(Document document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            document.Close(WdSaveOptions.wdDoNotSaveChanges);
        }

        /// <summary>
        /// Closes Word application without saving any changes.
        /// </summary>
        public static void ExitWordApplication()
        {
            Application app = GetWordApplication();
            app.Quit(WdSaveOptions.wdDoNotSaveChanges);
            Runtime.InteropServices.Marshal.ReleaseComObject(app);
            GC.Collect();
        }

        /// <summary>
        /// Gets existing Word application or initializes a new one.
        /// </summary>
        /// <returns>Word application.</returns>
        public static Application GetWordApplication()
        {
            return instance ?? (instance = new Application
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
                ScreenUpdating = false,
            });
        }

        /// <summary>
        /// Opens document in Word application and makes document active.
        /// </summary>
        /// <param name="filePath">Path to document file.</param>
        /// <returns>Document representation.</returns>
        public static Document OpenWordDocument(string filePath)
        {
            Document document = GetWordApplication().Documents.Open(filePath);
            document.Activate();
            return document;
        }
    }
}