using System.Collections.Generic;

namespace SmartOCR
{
    /// <summary>
    /// Class for processing parameters from start form.
    /// </summary>
    internal class WFProcess : IProcess
    {
        /// <summary>
        /// Collection of files to parse.
        /// </summary>
        private readonly List<string> files = new List<string>();

        /// <summary>
        /// Path to external config file.
        /// </summary>
        private readonly string config_file;

        /// <summary>
        /// Path to output workbook.
        /// </summary>
        private readonly string output_file;

        /// <summary>
        /// Initializes class instance with values from provided form.
        /// </summary>
        /// <param name="form">Instance of start form.</param>
        public WFProcess(StartForm form)
        {
            files = form.found_files;
            config_file = form.config_file;
            output_file = form.output_file;
        }

        public void ExecuteProcessing()
        {
            using (ParseEntryPoint entryPoint = new ParseEntryPoint(files, config_file, output_file))
            {
                entryPoint.TryGetData();
            }
        }
    }
}