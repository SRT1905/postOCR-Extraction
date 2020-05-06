namespace SmartOCR.Main
{
    using System;
    using System.Linq;
    using System.Windows.Forms;
    using SmartOCR.UI;
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

            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                args = GetArgumentsFromUI();
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

        private static string[] GetArgumentsFromUI()
        {
            StartForm startForm = new StartForm();
            if (startForm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return new string[] { };
            }

            return startForm.CmdArguments;
        }

        private static string[] CheckDebugEnablement(string[] args)
        {
            if (args.Length == 0)
            {
                return args;
            }

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