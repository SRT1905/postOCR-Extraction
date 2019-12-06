using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SmartOCR
{
    internal class Program
    {
        private static readonly int minimal_number_of_args = 2;
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

        private static List<string> GetFiles(IEnumerable<string> args)
        {
            return new List<string>(args);
        }

        private static List<string> GetFilesFromArgs(IEnumerable<string> args, string path_type)
        {
            if (path_type == "directory")
            {
                return GetFilesFromDirectories(args);
            }
            return GetFiles(args);
        }
        private static List<string> GetFilesFromDirectories(IEnumerable<string> args)
        {
            List<string> directories = new List<string>();
            foreach (var item in args)
            {
                directories.AddRange(Directory.EnumerateFiles(item).Where(file => !file.Contains("~")));
            }
            return directories;
        }

        private static void Main(string[] args)
        {
            if (args.Length >= minimal_number_of_args)
            {
                ExecuteProcessing(args);
            }
            else
            {
                PrintInvalidInputMessage();
            }
        }

        private static void PrintInvalidInputMessage()
        {
            System.Console.WriteLine("Enter valid document type and path(s) to file/directory.");
            System.Console.ReadKey();
        }

        private static string ValidateDocumentType(string doc_type)
        {
            HashSet<string> doc_types = new HashSet<string>
            {
                "Invoice"
            };
            return doc_types.FirstOrDefault(item => item.Contains(doc_type));
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