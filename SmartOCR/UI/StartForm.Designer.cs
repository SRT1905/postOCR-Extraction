namespace SmartOCR.UI
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
            this.labelConfig = new System.Windows.Forms.Label();
            this.textBoxConfig = new System.Windows.Forms.TextBox();
            this.buttonConfig = new System.Windows.Forms.Button();
            this.labelFiles = new System.Windows.Forms.Label();
            this.textBoxFiles = new System.Windows.Forms.TextBox();
            this.buttonFiles = new System.Windows.Forms.Button();
            this.buttonLaunch = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelConfig
            // 
            this.labelConfig.AutoSize = true;
            this.labelConfig.Location = new System.Drawing.Point(10, 14);
            this.labelConfig.Name = "labelConfig";
            this.labelConfig.Size = new System.Drawing.Size(88, 13);
            this.labelConfig.TabIndex = 0;
            this.labelConfig.Text = "Select config file:";
            // 
            // textBoxConfig
            // 
            this.textBoxConfig.Location = new System.Drawing.Point(13, 30);
            this.textBoxConfig.Name = "textBoxConfig";
            this.textBoxConfig.Size = new System.Drawing.Size(295, 20);
            this.textBoxConfig.TabIndex = 1;
            this.textBoxConfig.TextChanged += new System.EventHandler(this.TextBoxConfig_TextChanged);
            // 
            // buttonConfig
            // 
            this.buttonConfig.Location = new System.Drawing.Point(314, 28);
            this.buttonConfig.Name = "buttonConfig";
            this.buttonConfig.Size = new System.Drawing.Size(75, 22);
            this.buttonConfig.TabIndex = 2;
            this.buttonConfig.Text = "Browse";
            this.buttonConfig.UseVisualStyleBackColor = true;
            this.buttonConfig.Click += new System.EventHandler(this.ButtonConfig_Click);
            // 
            // labelFiles
            // 
            this.labelFiles.AutoSize = true;
            this.labelFiles.Location = new System.Drawing.Point(10, 57);
            this.labelFiles.Name = "labelFiles";
            this.labelFiles.Size = new System.Drawing.Size(113, 13);
            this.labelFiles.TabIndex = 3;
            this.labelFiles.Text = "Select files to process:";
            // 
            // textBoxFiles
            // 
            this.textBoxFiles.Location = new System.Drawing.Point(13, 73);
            this.textBoxFiles.Multiline = true;
            this.textBoxFiles.Name = "textBoxFiles";
            this.textBoxFiles.Size = new System.Drawing.Size(295, 20);
            this.textBoxFiles.TabIndex = 4;
            this.textBoxFiles.TextChanged += new System.EventHandler(this.TextBoxFiles_TextChanged);
            // 
            // buttonFiles
            // 
            this.buttonFiles.Location = new System.Drawing.Point(314, 71);
            this.buttonFiles.Name = "buttonFiles";
            this.buttonFiles.Size = new System.Drawing.Size(75, 22);
            this.buttonFiles.TabIndex = 5;
            this.buttonFiles.Text = "Browse";
            this.buttonFiles.UseVisualStyleBackColor = true;
            this.buttonFiles.Click += new System.EventHandler(this.ButtonFiles_Click);
            // 
            // buttonLaunch
            // 
            this.buttonLaunch.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonLaunch.Enabled = false;
            this.buttonLaunch.Location = new System.Drawing.Point(162, 109);
            this.buttonLaunch.Name = "buttonLaunch";
            this.buttonLaunch.Size = new System.Drawing.Size(75, 23);
            this.buttonLaunch.TabIndex = 6;
            this.buttonLaunch.Text = "Launch";
            this.buttonLaunch.UseVisualStyleBackColor = true;
            this.buttonLaunch.Click += new System.EventHandler(this.ButtonLaunch_Click);
            // 
            // StartForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(401, 154);
            this.Controls.Add(this.buttonLaunch);
            this.Controls.Add(this.buttonFiles);
            this.Controls.Add(this.textBoxFiles);
            this.Controls.Add(this.labelFiles);
            this.Controls.Add(this.buttonConfig);
            this.Controls.Add(this.textBoxConfig);
            this.Controls.Add(this.labelConfig);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "StartForm";
            this.Text = "Tool launch";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelConfig;
        private System.Windows.Forms.TextBox textBoxConfig;
        private System.Windows.Forms.Button buttonConfig;
        private System.Windows.Forms.Label labelFiles;
        private System.Windows.Forms.TextBox textBoxFiles;
        private System.Windows.Forms.Button buttonFiles;
        private System.Windows.Forms.Button buttonLaunch;
    }
}