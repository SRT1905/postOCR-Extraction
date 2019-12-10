using System.Collections.Generic;

namespace SmartOCR
{
    internal class WFProcess : IProcess
    {
        private List<string> Files { get; }
        private string ConfigFile { get; }
        private string DocumentType { get; }
        private string OutputFile { get; }

        public WFProcess()
        {
            Files = new List<string>();
        }

        public WFProcess(StartForm form)
        {
            this.Files = form.found_files;
            this.ConfigFile = form.config_file;
            this.DocumentType = form.document_type;
            this.OutputFile = form.output_file;
        }

        public void ExecuteProcessing()
        {
            using (ParseEntryPoint entryPoint = new ParseEntryPoint(DocumentType, Files, ConfigFile, OutputFile))
            {
                entryPoint.TryGetData();
            }
        }
    }
}