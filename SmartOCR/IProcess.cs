namespace SmartOCR
{
    /// <summary>
    /// Provides mechanism for initial input processing.
    /// </summary>
    public interface IProcess

    {       /// <summary>
            /// Calls for processing of identified documents.
            /// </summary>
        void ExecuteProcessing();
    }
}