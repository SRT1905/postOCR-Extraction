namespace SmartOCR.UI
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;

    /// <summary>
    /// Represents UI for tool arguments input.
    /// </summary>
    public partial class StartForm : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StartForm"/> class.
        /// </summary>
        public StartForm()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Gets arguments for tool launch.
        /// </summary>
        public string[] CmdArguments { get; private set; }

        private void ButtonConfig_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.Multiselect = false;
                dialog.Title = "Select config file";
                dialog.RestoreDirectory = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    this.textBoxConfig.Text = dialog.FileName;
                }
            }
        }

        private void ButtonFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Word documents (*.docx)|*.docx|All files (*.*)|*.*";
                dialog.Multiselect = true;
                dialog.Title = "Select files to process";
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    this.textBoxFiles.Text = string.Join(Environment.NewLine, dialog.FileNames);
                }
            }
        }

        private void ButtonLaunch_Click(object sender, EventArgs e)
        {
            string[] filesToProcess = this.textBoxFiles.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            List<string> argsAsList = new List<string>
            {
                this.textBoxConfig.Text,
            };

            argsAsList.AddRange(filesToProcess);
            this.CmdArguments = argsAsList.ToArray();
        }

        private void TextBoxConfig_TextChanged(object sender, EventArgs e)
        {
            this.TryEnableLaunchButton();
        }

        private void TryEnableLaunchButton()
        {
            this.buttonLaunch.Enabled = File.Exists(this.textBoxConfig.Text) &&
                                        this.textBoxFiles.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).All(item => File.Exists(item)) &&
                                        !string.IsNullOrEmpty(this.textBoxConfig.Text) &&
                                        !string.IsNullOrEmpty(this.textBoxFiles.Text);
        }

        private void TextBoxFiles_TextChanged(object sender, EventArgs e)
        {
            this.TryEnableLaunchButton();
        }
    }
}
