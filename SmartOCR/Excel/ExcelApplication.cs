﻿using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace SmartOCR
{
    /// <summary>
    /// Performs communication with MS Excel application.
    /// </summary>
    public static class ExcelApplication
    {
        /// <summary>
        /// Gets existing Excel application or initializes a new one.
        /// </summary>
        /// <returns>Excel application.</returns>
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

        /// <summary>
        /// Closes Excel application without saving any changes.
        /// </summary>
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

        /// <summary>
        /// Gets workbook in Excel application.
        /// </summary>
        /// <param name="path">Path to workbook file.</param>
        /// <returns>Workbook representation.</returns>
        public static Workbook OpenExcelWorkbook(string path)
        {
            Application app = GetExcelApplication();
            return TryGetWorkbook(app.Workbooks, path);
        }

        /// <summary>
        /// Gets active workbook in Excel application, closes it 
        /// and tries to delete the file, if its location is in Temp directory.
        /// </summary>
        public static void CloseActiveExcelWorkbook()
        {
            Application app = GetExcelApplication();
            app.ActiveWorkbook.Close(XlSaveAction.xlDoNotSaveChanges);
            TryDeleteTempFile(app.ActiveWorkbook.Path);
        }

        /// <summary>
        /// Check if file by provided path is in Temp directory.
        /// </summary>
        /// <param name="path">File path to check.</param>
        private static void TryDeleteTempFile(string path)
        {
            if (Path.GetDirectoryName(path) == Path.GetTempPath())
            {
                File.Delete(path);
            }
        }

        /// <summary>
        /// Tries to get <see cref="Workbook"/> by file path.
        /// If no <see cref="Workbook"/> is found,
        /// then application opens <see cref="Workbook"/> by path.
        /// </summary>
        /// <param name="workbooks">Collection of <see cref="Workbook"/> 
        /// objects in application.</param>
        /// <param name="path">File path to check.</param>
        /// <returns>Found or opened <see cref="Workbook"/> object.</returns>
        private static Workbook TryGetWorkbook(Workbooks workbooks, string path)
        {
            try
            {
                return workbooks[path];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return workbooks.Open(path);
            }
        }
    }
}