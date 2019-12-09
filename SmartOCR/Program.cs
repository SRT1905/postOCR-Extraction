using System;
using System.Windows.Forms;

namespace SmartOCR
{
    internal class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                using (StartForm form = new StartForm())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        new WFProcess(form).ExecuteProcessing();
                    }
                }
            }
            else
            {
                if (args.Length >= 2)
                {
                    var cmdProcessor = new CMDProcess(args);
                    if (cmdProcessor.ReadyToProcess)
                    {
                        cmdProcessor.ExecuteProcessing();
                    }
                }
                else
                {
                    Utilities.PrintInvalidInputMessage();
                }
            }
        }
    }
}