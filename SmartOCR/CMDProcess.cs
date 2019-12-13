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
        #region Fields
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
        /// Title of document type.
        /// </summary>
        private readonly string document_type;
        /// <summary>
        /// Specification of entered arguments.
        /// </summary>
        private readonly PathType entered_path_type;
        /// <summary>
        /// Collection of entered command propmpt arguments.
        /// </summary>
        private readonly List<string> _args;
        #endregion

        #region Properties
        /// <summary>
        /// Indicates whether command prompt arguments were successfully defined.
        /// </summary>
        public bool ReadyToProcess { get; }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes CMD processing object: identifies document type, type of entered paths and files by those paths.
        /// </summary>
        /// <param name="args">Array of command prompt arguments.</param>
        public CMDProcess(string[] args)
        {
            try
            {
                document_type = ValidateDocumentType(args[0]);
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
        #endregion

        #region Public methods
        /// <summary>
        /// Calls for processing of identified documents.
        /// </summary>
        public void ExecuteProcessing()
        {
            if (document_type == null || entered_path_type == PathType.None)
            {
                Utilities.PrintInvalidInputMessage();
                return;
            }
            using (var entryPoint = new ParseEntryPoint(document_type, GetFilesFromArgs()))
            {
                entryPoint.TryGetData();
            }
        }
        #endregion

        #region Private methods
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
            return new List<string>(_args);
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
                directories.AddRange(Directory.GetFiles(single_directory).Where(item => !item.StartsWith("~") && item.EndsWith(".docx")));
            }
            return directories;
        }
        /// <summary>
        /// Checks whether provided document type is in compliance with supportable document types.
        /// </summary>
        /// <param name="doc_type">Document type, entered from command propmt.</param>
        /// <returns>String representation of internally defined document type.</returns>
        private string ValidateDocumentType(string doc_type)
        {
            return Utilities.valid_document_types.FirstOrDefault(item => item.Contains(doc_type));
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
        #endregion

    }
}