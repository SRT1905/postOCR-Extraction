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
        private List<string> enteredArguments;

        /// <summary>
        /// Specification of entered arguments.
        /// </summary>
        private PathType enteredPathType;

        /// <summary>
        /// Path to external config file.
        /// </summary>
        private string configFile;

        /// <summary>
        /// Initializes a new instance of the <see cref="CMDProcess"/> class.
        /// Instance identifies type of entered paths and files by those paths.
        /// </summary>
        /// <param name="args">Array of command prompt arguments.</param>
        public CMDProcess(string[] args)
        {
            try
            {
                Utilities.Debug("Validating provided arguments.");
                if (!this.IsFirstArgumentValid(args))
                {
                    return;
                }

                this.InitializeFields(args);
            }
            catch (IndexOutOfRangeException)
            {
                Utilities.Debug(Properties.Resources.invalidInputMessage);
                this.IsReadyToProcess = false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether command prompt arguments were successfully defined.
        /// </summary>
        public bool IsReadyToProcess { get; private set; }

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

            string outputFile = this.GetOutputFilePath();
            Utilities.Debug($"Output file location: '{outputFile}'.");
            WaitForOutputFileDeletion(outputFile);
            using (var entryPoint = new ParseEntryPoint(this.GetFilesFromArgs(), this.configFile, outputFile))
            {
                entryPoint.TryGetData();
            }
        }

        private static void WaitForOutputFileDeletion(string outputFile)
        {
            while (File.Exists(outputFile))
            {
                Utilities.Debug("Waiting for existing output file to be deleted.");
                System.Threading.Thread.Sleep(1000);
            }
        }

        private static IEnumerable<string> GetFilesFromSingleDirectory(string singleDirectory)
        {
            return Directory.GetFiles(singleDirectory)
                            .Where(FilterInvalidFiles());
        }

        private static string GetConfigFileFromArgument(string argument)
        {
            if (Directory.Exists(argument))
            {
                return GetConfigFileFromDirectory(argument);
            }

            if (File.Exists(argument))
            {
                Utilities.Debug($"Found configuration file: '{argument}'.", 1);
                return argument;
            }

            return null;
        }

        private static string GetConfigFileFromDirectory(string argument)
        {
            Utilities.Debug($"Looking for configuration file in folder '{argument}'.", 1);
            string[] files = Directory.GetFiles(argument, "*.xlsx", SearchOption.TopDirectoryOnly);
            if (files.Length != 0)
            {
                Utilities.Debug($"Found configuration file: '{files[0]}'.", 1);
                return files[0];
            }

            return null;
        }

        private static Func<string, bool> FilterInvalidFiles()
        {
            return item => !Path.GetFileName(item).StartsWith("~", StringComparison.InvariantCultureIgnoreCase) &&
                           item.EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase);
        }

        private string GetOutputFilePath()
        {
            return this.enteredPathType == PathType.Directory
                ? Path.Combine(this.enteredArguments[0], "output.xlsx")
                : Path.Combine(Path.GetDirectoryName(this.enteredArguments[0]), "output.xlsx");
        }

        private void InitializeFields(string[] args)
        {
            this.enteredPathType = this.ValidatePath(args[1]);
            this.enteredArguments = args.Skip(1).ToList();
            this.IsReadyToProcess = true;
        }

        private bool IsFirstArgumentValid(string[] arguments)
        {
            if (arguments == null)
            {
                Utilities.Debug(Properties.Resources.invalidInputMessage);
                this.IsReadyToProcess = false;
                return false;
            }

            this.configFile = GetConfigFileFromArgument(arguments[0]);
            if (string.IsNullOrEmpty(this.configFile))
            {
                Utilities.Debug(Properties.Resources.noConfigFileFound);
                this.IsReadyToProcess = false;
                return false;
            }

            return true;
        }

        private List<string> GetFilesFromArgs()
        {
            return this.enteredPathType == PathType.Directory
                ? this.GetFilesFromDirectories()
                : this.enteredArguments
                    .Where(FilterInvalidFiles())
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
            foreach (string singleDirectory in this.enteredArguments)
            {
                directories.AddRange(GetFilesFromSingleDirectory(singleDirectory));
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
                Utilities.Debug("Arguments that represent documents to process will be processed as directories.", 1);
                return PathType.Directory;
            }

            if (File.Exists(path))
            {
                Utilities.Debug("Arguments that represent documents to process will be processed as files.", 1);
                return PathType.File;
            }

            Utilities.Debug("Arguments are identified neither as directories nor as files.", 1);
            return PathType.None;
        }
    }
}