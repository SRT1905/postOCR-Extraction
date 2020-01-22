using System;
using System.Windows.Forms;

namespace SmartOCR
{
    internal class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Utilities.PrintInvalidInputMessage();
            }
            else
            {
                var cmdProcessor = new CMDProcess(args);
                if (cmdProcessor.ReadyToProcess)
                {
                    cmdProcessor.ExecuteProcessing();
                }
            }
        }
    }
}