namespace SmartOCR
{
    using System;

    /// <summary>
    /// Entry point.
    /// </summary>
    internal class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Utilities.Debug(Properties.Resources.invalidInputMessage);
            }
            else
            {
                ProcessArguments(args);
            }
        }

        private static void ProcessArguments(string[] args)
        {
            var cmdProcessor = new CMDProcess(args);
            if (cmdProcessor.ReadyToProcess)
            {
                cmdProcessor.ExecuteProcessing();
            }
        }
    }
}