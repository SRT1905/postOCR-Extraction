using System.Collections.Generic;
using System.IO;

namespace SmartOCR
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                System.Console.WriteLine("Please enter document type and path(s) to file/directory.");
                System.Console.ReadKey();
                return;
            }
            string doc_type = ValidateDocumentType(args[0]);
            string path_type = ValidatePath(args[1]);
            if (doc_type == null || path_type == null)
            {
                System.Console.WriteLine("Enter valid document type and path(s) to file/directory.");
                System.Console.ReadKey();
                return;
            }
            List<string> files = new List<string>();
            if (path_type == "directory")
            {
                for (int i = 1; i < args.Length; i++)
                {
                    files.AddRange(Directory.EnumerateFiles(args[i]));
                }
                files.RemoveAll(item => item.Contains("~"));
            }

            if (path_type == "file")
            {
                for (int i = 1; i < args.Length; i++)
                {
                    files.Add(args[i]);
                }
            }

            _ = new ParseEntryPoint(doc_type, files);
        }
        static string ValidateDocumentType(string doc_type)
        {
            HashSet<string> doc_types = new HashSet<string>
            {
                "Invoice"
            };
            foreach (string item in doc_types)
            {
                if (item.Contains(doc_type))
                {
                    return item;
                }
            }
            return null;
        }
        static string ValidatePath(string path)
        {
            if (Directory.Exists(path))
            {
                return "directory";
            }

            if (File.Exists(path))
            {
                return "file";
            }

            return null;
        }
    }
}
