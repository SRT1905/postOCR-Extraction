namespace SmartOCR
{
    using System;
    using System.Linq;

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
            if (args[0].StartsWith("-d"))
            {
                args = ProcessDebugParameter(args);
            }

            return args;
        }

        private static string[] ProcessDebugParameter(string[] args)
        {
            Utilities.EnableDebug = true;
            SetDebugLevel(args[0]);
            args = RemoveDebugParameter(args);
            return args;
        }

        private static string[] RemoveDebugParameter(string[] args)
        {
            return args.Skip(1).ToArray();
        }

        private static void ProcessArguments(string[] args)
        {
            var cmdProcessor = new CMDProcess(args);
            if (cmdProcessor.IsReadyToProcess)
            {
                cmdProcessor.ExecuteProcessing();
                Utilities.Debug($"Processing done!");
            }
        }

        private static void SetDebugLevel(string argument)
        {
            argument = argument.Replace("-d", string.Empty);
            if (argument.Length != 0)
            {
                Utilities.DebugLevel = Convert.ToInt32(argument);
            }
        }
    }
}