using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SmartOCR
{
    internal class CommandLineStart
    {
        private static string doc_type;
        private static string path_type;

        private static void ExecuteProcessing(string[] args)
        {
            doc_type = ValidateDocumentType(args[0]);
            path_type = ValidatePath(args[1]);
            if (doc_type == null || path_type == null)
            {
                PrintInvalidInputMessage();
                return;
            }
            using (ParseEntryPoint entryPoint = new ParseEntryPoint(doc_type, GetFilesFromArgs(args.Skip(1), path_type)))
            {
                entryPoint.TryGetData();
            }
        }
        private static List<string> GetFilesFromArgs(IEnumerable<string> args, string path_type)
        {
            if (path_type == "directory")
            {
                return GetFilesFromDirectories(args);
            }
            return new List<string>(args);
        }
        private static List<string> GetFilesFromDirectories(IEnumerable<string> args)
        {
            List<string> directories = new List<string>();
            foreach (var item in args)
            {
                directories.AddRange(Directory.GetFiles(item).Where(file => !file.Contains("~")));
            }
            return directories;
        }

        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                System.Windows.Forms.Application.EnableVisualStyles();
                StartForm form = new StartForm();
                _ = form.ShowDialog();
            }
            else
            {
                if (args.Length >= 2)
                {
                    ExecuteProcessing(args);
                }
                else
                {
                    PrintInvalidInputMessage();
                }
            }

        }

        private static void PrintInvalidInputMessage()
        {
            System.Console.WriteLine("Enter valid document type and path(s) to file/directory.");
            System.Console.ReadKey();
        }

        private static string ValidateDocumentType(string doc_type)
        {
            
            return Utilities.valid_document_types.FirstOrDefault(item => item.Contains(doc_type));
        }

        private static string ValidatePath(string path)
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