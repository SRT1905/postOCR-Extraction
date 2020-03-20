namespace SmartOCR.Main
{
    /// <summary>
    /// Specifies the type of entered command prompt arguments.
    /// </summary>
    public enum PathType
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
}