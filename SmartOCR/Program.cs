﻿using System;
using System.Collections.Generic;

namespace SmartOCR
{
    internal class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length < 2)
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