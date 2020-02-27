namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    /// <summary>
    /// Specifies the type of entered command propmt arguments.
    /// </summary>
    internal enum PathType : int
    {
        /// <summary>
        /// Represents invalid path type.
        /// </summary>
        None,

        /// <summary>
        /// Represents path as a directory.
        /// </summary>
        Directory,

        /// <summary>
        /// Represents path as a file.
        /// </summary>
        File,
    }

    /// <summary>
    /// CMDProcess is used to process data from command prompt: define document type and files to parse.
    /// </summary>
    public class CMDProcess
    {
        /// <summary>
        /// Collection of entered command prompt arguments.
        /// </summary>
        private readonly List<string> enteredArguments;

        /// <summary>
        /// Specification of entered arguments.
        /// </summary>
        private readonly PathType enteredPathType;

        /// <summary>
        /// Path to external config file.
        /// </summary>
        private readonly string configFile;

        /// <summary>
        /// Initializes a new instance of the <see cref="CMDProcess"/> class.
        /// Instance identifies document type, type of entered paths and files by those paths.
        /// </summary>
        /// <param name="args">Array of command prompt arguments.</param>
        public CMDProcess(string[] args)
        {
            try
            {
                if (args == null)
                {
                    Utilities.Debug(Properties.Resources.invalidInputMessage);
                    this.ReadyToProcess = false;
                    return;
                }

                if (Directory.Exists(args[0]))
                {
                    string[] files = Directory.GetFiles(args[0], "*.xlsx", SearchOption.TopDirectoryOnly);
                    if (files.Length != 0)
                    {
                        this.configFile = files[0];
                    }
                    else
                    {
                        Console.WriteLine(Properties.Resources.noConfigFileFound);
                        return;
                    }
                }
                else if (File.Exists(args[0]))
                {
                    this.configFile = args[0];
                }

                this.enteredPathType = this.ValidatePath(args[1]);
                this.enteredArguments = args.Skip(1).ToList();
                this.ReadyToProcess = true;
            }
            catch (IndexOutOfRangeException)
            {
                Utilities.Debug(Properties.Resources.invalidInputMessage);
                this.ReadyToProcess = false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether command prompt arguments were successfully defined.
        /// </summary>
        public bool ReadyToProcess { get; }

        /// <summary>
        /// Sets output file and reads data from provided files.
        /// </summary>
        public void ExecuteProcessing()
        {
            if (this.enteredPathType == PathType.None || this.configFile == null)
            {
                Utilities.Debug(Properties.Resources.invalidInputMessage);
                return;
            }

            string outputFile = this.enteredPathType == PathType.Directory
                ? Path.Combine(this.enteredArguments[0], "output.xlsx")
                : Path.Combine(Path.GetDirectoryName(this.enteredArguments[0]), "output.xlsx");

            while (File.Exists(outputFile))
            {
                continue;
            }

            using (var entryPoint = new ParseEntryPoint(this.GetFilesFromArgs(), this.configFile, outputFile))
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
            if (this.enteredPathType == PathType.Directory)
            {
                return this.GetFilesFromDirectories();
            }

            return new List<string>(this.enteredArguments)
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
            for (int i = 0; i < this.enteredArguments.Count; i++)
            {
                string singleDirectory = this.enteredArguments[i];
                directories.AddRange(Directory.GetFiles(singleDirectory).Where(item => !Path.GetFileName(item).StartsWith("~", StringComparison.InvariantCultureIgnoreCase) && item.EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase)));
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