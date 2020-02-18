using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// CMDProcess is used to process data from command prompt: define document type and files to parse.
    /// </summary>
    public class CMDProcess : IProcess
    {

        /// <summary>
        /// Specifies the type of entered command propmt arguments
        /// </summary>
        private enum PathType : int
        {
            None,
            Directory,
            File
        }

        /// <summary>
        /// Specification of entered arguments.
        /// </summary>
        private readonly PathType entered_path_type;

        /// <summary>
        /// Collection of entered command propmpt arguments.
        /// </summary>
        private readonly List<string> _args;

        /// <summary>
        /// Path to external config file.
        /// </summary>
        private readonly string config_file;

        /// <summary>
        /// Indicates whether command prompt arguments were successfully defined.
        /// </summary>
        public bool ReadyToProcess { get; }

        /// <summary>
        /// Initializes CMD processing object: identifies document type, type of entered paths and files by those paths.
        /// </summary>
        /// <param name="args">Array of command prompt arguments.</param>
        public CMDProcess(string[] args)
        {
            try
            {
                if (args == null)
                {
                    Utilities.PrintInvalidInputMessage();
                    ReadyToProcess = false;
                    return;
                }
                if (Directory.Exists(args[0]))
                {
                    string[] files = Directory.GetFiles(args[0], "*.xlsx", SearchOption.TopDirectoryOnly);
                    if (files.Length != 0)
                    {
                        config_file = files[0];
                    }
                    else
                    {
                        Console.WriteLine(Properties.Resources.noConfigFileFound);
                        return;
                    }
                }
                else if (File.Exists(args[0]))
                {
                    config_file = args[0];
                }
                entered_path_type = ValidatePath(args[1]);
                _args = args.Skip(1).ToList();
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
            if (entered_path_type == PathType.None || config_file == null)
            {
                Utilities.PrintInvalidInputMessage();
                return;
            }

            string output_file;
            if (entered_path_type == PathType.Directory)
            {
                output_file = Path.Combine(_args[0], "output.xlsx");
            }
            else
            {
                output_file = Path.Combine(Path.GetDirectoryName(_args[0]), "output.xlsx");
            }

            while (File.Exists(output_file))
            {
                continue;
            }
            using (var entryPoint = new ParseEntryPoint(GetFilesFromArgs(), config_file, output_file))
            {
                entryPoint.TryGetData();
            }
        }

        /// <summary>
        /// Processes entered path type and path argument(s).
        /// </summary>
        /// <returns>List of files, specified by arguments.</returns>
        private List<string> GetFilesFromArgs()
        {
            if (entered_path_type == PathType.Directory)
            {
                return GetFilesFromDirectories();
            }
            return new List<string>(_args)
                .Where(item => !Path.GetFileName(item).StartsWith("~", StringComparison.InvariantCultureIgnoreCase) && item.EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase))
                .ToList();
        }

        /// <summary>
        /// Gets files from directories, provided in Args field.
        /// Files with .docx extension are taken.
        /// Files with '~' symbol in title are ignored.
        /// </summary>
        /// <returns>Collection of files in provided directories.</returns>
        private List<string> GetFilesFromDirectories()
        {
            List<string> directories = new List<string>();
            for (int i = 0; i < _args.Count; i++)
            {
                string single_directory = _args[i];
                directories.AddRange(Directory.GetFiles(single_directory).Where(item => !Path.GetFileName(item).StartsWith("~", StringComparison.InvariantCultureIgnoreCase) && item.EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase)));
            }
            return directories;
        }

        /// <summary>
        /// Checks whether argument, provided after document type is path to directory or file.
        /// </summary>
        /// <param name="path">Path, entered from command prompt.</param>
        /// <returns>PathType enumeration item, that indicates path type.</returns>
        private PathType ValidatePath(string path)
        {
            if (Directory.Exists(path))
            {
                return PathType.Directory;
            }
            if (File.Exists(path))
            {
                return PathType.File;
            }
            return PathType.None;
        }
    }
}