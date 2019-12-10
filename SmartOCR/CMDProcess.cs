using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SmartOCR
{
    public class CMDProcess : IProcess
    {
        private string DocumentType { get; }
        private string PathType { get; }
        private IEnumerable<string> Args { get; }
        public bool ReadyToProcess { get; }

        public CMDProcess(string[] args)
        {
            try
            {
                DocumentType = ValidateDocumentType(args[0]);
                PathType = ValidatePath(args[1]);
                this.Args = args.Skip(1);
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
            if (DocumentType == null || PathType == null)
            {
                Utilities.PrintInvalidInputMessage();
                return;
            }
            using (var entryPoint = new ParseEntryPoint(DocumentType, GetFilesFromArgs()))
            {
                entryPoint.TryGetData();
            }
        }

        private List<string> GetFilesFromArgs()
        {
            if (PathType == "directory")
            {
                return GetFilesFromDirectories();
            }
            return new List<string>(Args);
        }

        private List<string> GetFilesFromDirectories()
        {
            List<string> directories = new List<string>();
            foreach (var item in Args)
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