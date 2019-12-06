namespace SmartOCR
{
    partial class StartForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label_input_config = new System.Windows.Forms.Label();
            this.textbox_input_config = new System.Windows.Forms.TextBox();
            this.button_config = new System.Windows.Forms.Button();
            this.label_doc_types = new System.Windows.Forms.Label();
            this.combobox_doc_types = new System.Windows.Forms.ComboBox();
            this.label_chargecode = new System.Windows.Forms.Label();
            this.textbox_chargecode = new System.Windows.Forms.TextBox();
            this.label_directory = new System.Windows.Forms.Label();
            this.textbox_directory = new System.Windows.Forms.TextBox();
            this.button_directory = new System.Windows.Forms.Button();
            this.label_file_specification = new System.Windows.Forms.Label();
            this.textbox_file_specification = new System.Windows.Forms.TextBox();
            this.checkbox_include_subdirectories = new System.Windows.Forms.CheckBox();
            this.button_run = new System.Windows.Forms.Button();
            this.errorprovider = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.errorprovider)).BeginInit();
            this.SuspendLayout();
            // 
            // label_input_config
            // 
            this.label_input_config.AutoSize = true;
            this.label_input_config.Location = new System.Drawing.Point(13, 13);
            this.label_input_config.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label_input_config.Name = "label_input_config";
            this.label_input_config.Size = new System.Drawing.Size(56, 13);
            this.label_input_config.TabIndex = 0;
            this.label_input_config.Text = "Config file:";
            // 
            // textbox_input_config
            // 
            this.textbox_input_config.Location = new System.Drawing.Point(16, 31);
            this.textbox_input_config.Margin = new System.Windows.Forms.Padding(5);
            this.textbox_input_config.Name = "textbox_input_config";
            this.textbox_input_config.Size = new System.Drawing.Size(282, 20);
            this.textbox_input_config.TabIndex = 1;
            this.textbox_input_config.TextChanged += new System.EventHandler(this.Textbox_input_config_TextChanged);
            // 
            // button_config
            // 
            this.button_config.Location = new System.Drawing.Point(308, 31);
            this.button_config.Margin = new System.Windows.Forms.Padding(5);
            this.button_config.Name = "button_config";
            this.button_config.Size = new System.Drawing.Size(83, 20);
            this.button_config.TabIndex = 2;
            this.button_config.Text = "Browse";
            this.button_config.UseVisualStyleBackColor = true;
            this.button_config.Click += new System.EventHandler(this.Button_config_Click);
            // 
            // label_doc_types
            // 
            this.label_doc_types.AutoSize = true;
            this.label_doc_types.Location = new System.Drawing.Point(13, 56);
            this.label_doc_types.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label_doc_types.Name = "label_doc_types";
            this.label_doc_types.Size = new System.Drawing.Size(82, 13);
            this.label_doc_types.TabIndex = 3;
            this.label_doc_types.Text = "Document type:";
            // 
            // combobox_doc_types
            // 
            this.combobox_doc_types.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.combobox_doc_types.FormattingEnabled = true;
            this.combobox_doc_types.Location = new System.Drawing.Point(16, 74);
            this.combobox_doc_types.Margin = new System.Windows.Forms.Padding(5);
            this.combobox_doc_types.Name = "combobox_doc_types";
            this.combobox_doc_types.Size = new System.Drawing.Size(180, 21);
            this.combobox_doc_types.Sorted = true;
            this.combobox_doc_types.TabIndex = 4;
            this.combobox_doc_types.SelectedIndexChanged += new System.EventHandler(this.Combobox_doc_types_SelectedIndexChanged);
            // 
            // label_chargecode
            // 
            this.label_chargecode.AutoSize = true;
            this.label_chargecode.Location = new System.Drawing.Point(208, 56);
            this.label_chargecode.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label_chargecode.Name = "label_chargecode";
            this.label_chargecode.Size = new System.Drawing.Size(97, 13);
            this.label_chargecode.TabIndex = 5;
            this.label_chargecode.Text = "Engagement code:";
            // 
            // textbox_chargecode
            // 
            this.textbox_chargecode.Location = new System.Drawing.Point(211, 75);
            this.textbox_chargecode.MaxLength = 8;
            this.textbox_chargecode.Name = "textbox_chargecode";
            this.textbox_chargecode.Size = new System.Drawing.Size(180, 20);
            this.textbox_chargecode.TabIndex = 6;
            this.textbox_chargecode.Validating += new System.ComponentModel.CancelEventHandler(this.Textbox_chargecode_Validating);
            this.textbox_chargecode.Validated += new System.EventHandler(this.Textbox_chargecode_Validated);
            // 
            // label_directory
            // 
            this.label_directory.AutoSize = true;
            this.label_directory.Location = new System.Drawing.Point(13, 100);
            this.label_directory.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label_directory.Name = "label_directory";
            this.label_directory.Size = new System.Drawing.Size(87, 13);
            this.label_directory.TabIndex = 7;
            this.label_directory.Text = "Search directory:";
            // 
            // textbox_directory
            // 
            this.textbox_directory.Location = new System.Drawing.Point(16, 118);
            this.textbox_directory.Margin = new System.Windows.Forms.Padding(5);
            this.textbox_directory.Name = "textbox_directory";
            this.textbox_directory.Size = new System.Drawing.Size(282, 20);
            this.textbox_directory.TabIndex = 8;
            this.textbox_directory.TextChanged += new System.EventHandler(this.Textbox_directory_TextChanged);
            // 
            // button_directory
            // 
            this.button_directory.Location = new System.Drawing.Point(308, 118);
            this.button_directory.Margin = new System.Windows.Forms.Padding(5);
            this.button_directory.Name = "button_directory";
            this.button_directory.Size = new System.Drawing.Size(83, 20);
            this.button_directory.TabIndex = 9;
            this.button_directory.Text = "Browse";
            this.button_directory.UseVisualStyleBackColor = true;
            this.button_directory.Click += new System.EventHandler(this.Button_directory_Click);
            // 
            // label_file_specification
            // 
            this.label_file_specification.AutoSize = true;
            this.label_file_specification.Location = new System.Drawing.Point(13, 143);
            this.label_file_specification.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label_file_specification.Name = "label_file_specification";
            this.label_file_specification.Size = new System.Drawing.Size(88, 13);
            this.label_file_specification.TabIndex = 10;
            this.label_file_specification.Text = "File specification:";
            // 
            // textbox_file_specification
            // 
            this.textbox_file_specification.Location = new System.Drawing.Point(16, 161);
            this.textbox_file_specification.Margin = new System.Windows.Forms.Padding(5);
            this.textbox_file_specification.Name = "textbox_file_specification";
            this.textbox_file_specification.Size = new System.Drawing.Size(282, 20);
            this.textbox_file_specification.TabIndex = 11;
            this.textbox_file_specification.Text = "*.*";
            this.textbox_file_specification.TextChanged += new System.EventHandler(this.Textbox_file_specification_TextChanged);
            // 
            // checkbox_include_subdirectories
            // 
            this.checkbox_include_subdirectories.AutoSize = true;
            this.checkbox_include_subdirectories.Location = new System.Drawing.Point(16, 191);
            this.checkbox_include_subdirectories.Margin = new System.Windows.Forms.Padding(5);
            this.checkbox_include_subdirectories.Name = "checkbox_include_subdirectories";
            this.checkbox_include_subdirectories.Size = new System.Drawing.Size(133, 17);
            this.checkbox_include_subdirectories.TabIndex = 12;
            this.checkbox_include_subdirectories.Text = "Include SubDirectories";
            this.checkbox_include_subdirectories.UseVisualStyleBackColor = true;
            // 
            // button_run
            // 
            this.button_run.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button_run.Enabled = false;
            this.button_run.Location = new System.Drawing.Point(168, 223);
            this.button_run.Margin = new System.Windows.Forms.Padding(5);
            this.button_run.Name = "button_run";
            this.button_run.Size = new System.Drawing.Size(75, 23);
            this.button_run.TabIndex = 13;
            this.button_run.Text = "Run";
            this.button_run.UseVisualStyleBackColor = true;
            this.button_run.Click += new System.EventHandler(this.Button_run_Click);
            // 
            // errorprovider
            // 
            this.errorprovider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorprovider.ContainerControl = this;
            // 
            // StartForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(409, 260);
            this.Controls.Add(this.button_run);
            this.Controls.Add(this.checkbox_include_subdirectories);
            this.Controls.Add(this.textbox_file_specification);
            this.Controls.Add(this.label_file_specification);
            this.Controls.Add(this.button_directory);
            this.Controls.Add(this.textbox_directory);
            this.Controls.Add(this.label_directory);
            this.Controls.Add(this.textbox_chargecode);
            this.Controls.Add(this.label_chargecode);
            this.Controls.Add(this.combobox_doc_types);
            this.Controls.Add(this.label_doc_types);
            this.Controls.Add(this.button_config);
            this.Controls.Add(this.textbox_input_config);
            this.Controls.Add(this.label_input_config);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "StartForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "SmartOCR";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.StartForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.errorprovider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label_input_config;
        private System.Windows.Forms.TextBox textbox_input_config;
        private System.Windows.Forms.Button button_config;
        private System.Windows.Forms.Label label_doc_types;
        private System.Windows.Forms.ComboBox combobox_doc_types;
        private System.Windows.Forms.Label label_chargecode;
        private System.Windows.Forms.TextBox textbox_chargecode;
        private System.Windows.Forms.Label label_directory;
        private System.Windows.Forms.TextBox textbox_directory;
        private System.Windows.Forms.Button button_directory;
        private System.Windows.Forms.Label label_file_specification;
        private System.Windows.Forms.TextBox textbox_file_specification;
        private System.Windows.Forms.CheckBox checkbox_include_subdirectories;
        private System.Windows.Forms.Button button_run;
        private System.Windows.Forms.ErrorProvider errorprovider;
    }
}