using System.Collections.Generic;

namespace SmartOCR
{
    internal class WFProcess : IProcess
    {
        private List<string> files { get; }
        private string config_file { get; }
        private string doc_type { get; }
        private string output_file { get; }

        public WFProcess()
        {
            files = new List<string>();
        }

        public WFProcess(StartForm form)
        {
            this.files = form.found_files;
            this.config_file = form.config_file;
            this.doc_type = form.document_type;
            this.output_file = form.output_file;
        }

        public void ExecuteProcessing()
        {
            using (ParseEntryPoint entryPoint = new ParseEntryPoint(doc_type, files, config_file, output_file))
            {
                entryPoint.TryGetData();
            }
        }
    }
}