namespace SmartOCR
{
    using System;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// Performs communication with MS Word application.
    /// </summary>
    public static class WordApplication
    {
        /// <summary>
        /// Single instance of Word application.
        /// </summary>
        private static Word.Application instance;

        /// <summary>
        /// Closes provided Word document without saving any changes.
        /// </summary>
        /// <param name="document">A Word document.</param>
        public static void CloseDocument(Word.Document document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        }

        /// <summary>
        /// Closes Word application without saving any changes.
        /// </summary>
        public static void ExitWordApplication()
        {
            Word.Application app = GetWordApplication();
            app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            if (app != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }

            GC.Collect();
        }

        /// <summary>
        /// Gets existing Word application or initializes a new one.
        /// </summary>
        /// <returns>Word application.</returns>
        public static Word.Application GetWordApplication()
        {
            if (instance == null)
            {
                instance = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                    ScreenUpdating = false,
                };
            }

            return instance;
        }

        /// <summary>
        /// Opens document in Word application and makes document active.
        /// </summary>
        /// <param name="filePath">Path to document file.</param>
        /// <returns>Document representation.</returns>
        public static Word.Document OpenWordDocument(string filePath)
        {
            Word.Document document = GetWordApplication().Documents.Open(filePath);
            document.Activate();
            return document;
        }
    }
}