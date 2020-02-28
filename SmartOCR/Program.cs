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
            args = CheckDebugEnablement(args);

            if (args.Length < 2)
            {
                Utilities.Debug(Properties.Resources.invalidInputMessage);
            }
            else
            {
                ProcessArguments(args);
            }
        }

        private static string[] CheckDebugEnablement(string[] args)
        {
            if (args[0] == "-d")
            {
                Utilities.EnableDebug = true;
                args = RemoveDebugParameter(args);
            }

            return args;
        }

        private static string[] RemoveDebugParameter(string[] args)
        {
            string[] newArgs = new string[args.Length - 1];
            for (int i = 1; i < args.Length; i++)
            {
                newArgs[i - 1] = args[i];
            }

            return newArgs;
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