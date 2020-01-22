using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;

namespace SmartOCR
{
    public partial class StartForm : Form
    {

        public string ConfigFile { get; set; }
        public string SearchDirectory { get; set; }
        public string ChargeCode { get; set; }
        public List<string> FoundFiles { get; } = new List<string>();
        public string SearchPattern { get; set; }
        public string DocumentType { get; set; }
        public string OutputFile { get; set; }
        private bool is_output_valid;
        public StartForm()
        {
            InitializeComponent();
            InitializeDocTypes();
        }

        private void InitializeDocTypes()
        {
            combobox_doc_types.Items.AddRange(Utilities.valid_document_types.ToArray());
        }
        private void Button_config_Click(object sender, System.EventArgs e)
        {
            string initial_directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = Properties.Resources.openFileDialogExcelFilter;
                dialog.Multiselect = false;
                dialog.InitialDirectory = initial_directory;
                dialog.RestoreDirectory = true;
                dialog.Title = Properties.Resources.openConfigFileDialogTitle;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ConfigFile = dialog.FileName;
                    textbox_input_config.Text = dialog.FileName;
                }
            }
        }

        private void Textbox_chargecode_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!ValidateChargeCode(textbox_chargecode.Text, out string error_msg))
            {
                e.Cancel = true;
                textbox_chargecode.Select(0, textbox_chargecode.Text.Length);
                errorprovider.SetError(textbox_chargecode, error_msg);
            }
        }

        private bool ValidateChargeCode(string charge_code, out string error_message)
        {
            error_message = "Enter valid engagement code.";
            return !new List<bool>()
            {
                charge_code.Length == 8,
                charge_code.All(char.IsDigit),
                new HashSet<char>(charge_code.ToCharArray()).Count > 1,
                charge_code != "12345678",
                charge_code != "87654321"
            }.Contains(false);
        }

        private void Textbox_chargecode_Validated(object sender, System.EventArgs e)
        {
            ChargeCode = textbox_chargecode.Text;
            errorprovider.SetError(textbox_chargecode, string.Empty);
            TryEnableRunButton();
        }

        private void Button_directory_Click(object sender, System.EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = Properties.Resources.folderBrowserDialogDescription;
                dialog.ShowNewFolderButton = true;
            
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    SearchDirectory = dialog.SelectedPath;
                    textbox_directory.Text = dialog.SelectedPath;
                }
            }
        }

        private bool TryEnableRunButton()
        {
            bool status = !new List<bool>()
            {
                !string.IsNullOrEmpty(textbox_chargecode.Text),
                Directory.Exists(textbox_directory.Text),
                File.Exists(textbox_input_config.Text),
                !string.IsNullOrEmpty(textbox_file_specification.Text),
                combobox_doc_types.SelectedIndex != -1,
            }.Contains(false);
            button_run.Enabled = status;
            return status;
        }

        private void Textbox_input_config_TextChanged(object sender, System.EventArgs e)
        {
            TryEnableRunButton();
        }

        private void Combobox_doc_types_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            TryEnableRunButton();
        }

        private void Textbox_directory_TextChanged(object sender, System.EventArgs e)
        {
            TryEnableRunButton();
        }

        private void Textbox_file_specification_TextChanged(object sender, System.EventArgs e)
        {
            TryEnableRunButton();
        }

        private void Button_run_Click(object sender, System.EventArgs e)
        {
            SearchPattern = textbox_file_specification.Text;
            FoundFiles.AddRange(GetFilesFromInput());
            DocumentType = combobox_doc_types.SelectedItem.ToString();
            is_output_valid = File.Exists(OutputFile) || string.IsNullOrEmpty(OutputFile);
        }

        private void StartForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (button_run.Enabled && TryEnableRunButton())
            {
                if (FoundFiles.Count == 0)
                {
                    e.Cancel = true;
                    errorprovider.SetError(button_run, "No files were found. Check input parameters.");
                }
                if (!is_output_valid)
                {
                    e.Cancel = true;
                    errorprovider.SetError(textbox_save_location, "Invalid output file location.");
                }
            }
        }

        private List<string> GetFilesFromInput()
        {
            SearchOption search_option = checkbox_include_subdirectories.Checked
                ? SearchOption.AllDirectories
                : SearchOption.TopDirectoryOnly;
            try
            {
                return Directory.EnumerateFiles(SearchDirectory, SearchPattern, search_option).ToList();
            }
            catch (ArgumentException)
            {
                return new List<string>();
            }
        }

        private void Button_save_location_Click(object sender, System.EventArgs e)
        {
            string initial_directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = Properties.Resources.openFileDialogExcelFilter;
                dialog.InitialDirectory = initial_directory;
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    OutputFile = dialog.FileName;
                    textbox_save_location.Text = dialog.FileName;
                }
            }
        }

        private void Textbox_save_location_TextChanged(object sender, System.EventArgs e)
        {
            TryEnableRunButton();
        }
    }
}