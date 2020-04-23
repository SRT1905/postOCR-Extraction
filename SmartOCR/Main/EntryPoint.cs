namespace SmartOCR.Main
{
    using System;
    using System.Linq;
    using Utilities = SmartOCR.Utilities.UtilitiesClass;

    /// <summary>
    /// Used as entry point for Word document processing.
    /// </summary>
    internal class EntryPoint
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (DateTime.Today >= new DateTime(2020, 6, 1))
            {
                Console.WriteLine("Test period has ended.");
                return;
            }

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
            var cmdProcessor = new CmdProcess(args);
            if (!cmdProcessor.IsReadyToProcess)
            {
                return;
            }

            cmdProcessor.ExecuteProcessing();
            Utilities.Debug("Processing done!");
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