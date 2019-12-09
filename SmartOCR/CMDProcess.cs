using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SmartOCR
{
    public class CMDProcess : IProcess
    {
        private string doc_type { get; }
        private string path_type { get; }
        private IEnumerable<string> args { get; }
        public bool ReadyToProcess { get; }

        public CMDProcess(string[] args)
        {
            try
            {
                doc_type = ValidateDocumentType(args[0]);
                path_type = ValidatePath(args[1]);
                this.args = args.Skip(1);
                ReadyToProcess = true;
            }
            catch (IndexOutOfRangeException)
            {
                Utilities.PrintInvalidInputMessage();
                ReadyToProcess = false;
            }
        }

        public void ExecuteProcessing()
        {
            if (doc_type == null || path_type == null)
            {
                Utilities.PrintInvalidInputMessage();
                return;
            }
            using (var entryPoint = new ParseEntryPoint(doc_type, GetFilesFromArgs()))
            {
                entryPoint.TryGetData();
            }
        }

        private List<string> GetFilesFromArgs()
        {
            if (path_type == "directory")
            {
                return GetFilesFromDirectories();
            }
            return new List<string>(args);
        }

        private List<string> GetFilesFromDirectories()
        {
            List<string> directories = new List<string>();
            foreach (var item in args)
            {
                directories.AddRange(Directory.GetFiles(item).Where(file => !file.Contains("~")));
            }
            return directories;
        }

        public string ValidateDocumentType(string doc_type)
        {
            return Utilities.valid_document_types.FirstOrDefault(item => item.Contains(doc_type));
        }

        public string ValidatePath(string path)
        {
            if (Directory.Exists(path))
            {
                return "directory";
            }
            else if (File.Exists(path))
            {
                return "file";
            }
            else
            {
                return null;
            }
        }
    }
}