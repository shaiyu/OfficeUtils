namespace OfficeUtils
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private System.Windows.Forms.Button btnSelectFiles;
        private System.Windows.Forms.Button btnSelectOutput;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ListBox lstFiles;
        private System.Windows.Forms.ListBox lstLog;
        private System.Windows.Forms.NumericUpDown numMaxRows;
        private System.Windows.Forms.NumericUpDown numHeaderRows;
        private System.Windows.Forms.Label lblMaxRows;
        private System.Windows.Forms.Label lblHeaderRows;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            this.btnSelectFiles = new System.Windows.Forms.Button();
            this.btnSelectOutput = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.lstFiles = new System.Windows.Forms.ListBox();
            this.lstLog = new System.Windows.Forms.ListBox();
            this.numMaxRows = new System.Windows.Forms.NumericUpDown();
            this.numHeaderRows = new System.Windows.Forms.NumericUpDown();
            this.lblMaxRows = new System.Windows.Forms.Label();
            this.lblHeaderRows = new System.Windows.Forms.Label();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.numMaxRows)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHeaderRows)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSelectFiles
            // 
            this.btnSelectFiles.Location = new System.Drawing.Point(12, 12);
            this.btnSelectFiles.Name = "btnSelectFiles";
            this.btnSelectFiles.Size = new System.Drawing.Size(150, 27);
            this.btnSelectFiles.TabIndex = 0;
            this.btnSelectFiles.Text = "选择 Excel 文件...";
            this.btnSelectFiles.UseVisualStyleBackColor = true;
            this.btnSelectFiles.Click += new System.EventHandler(this.btnSelectFiles_Click);
            // 
            // btnSelectOutput
            // 
            this.btnSelectOutput.Location = new System.Drawing.Point(168, 12);
            this.btnSelectOutput.Name = "btnSelectOutput";
            this.btnSelectOutput.Size = new System.Drawing.Size(150, 27);
            this.btnSelectOutput.TabIndex = 1;
            this.btnSelectOutput.Text = "选择输出目录...";
            this.btnSelectOutput.UseVisualStyleBackColor = true;
            this.btnSelectOutput.Click += new System.EventHandler(this.btnSelectOutput_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(324, 12);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(120, 27);
            this.btnStart.TabIndex = 2;
            this.btnStart.Text = "确定拆分";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // lblHeaderRows
            // 
            this.lblHeaderRows.AutoSize = true;
            this.lblHeaderRows.Location = new System.Drawing.Point(460, 17);
            this.lblHeaderRows.Name = "lblHeaderRows";
            this.lblHeaderRows.Size = new System.Drawing.Size(140, 15);
            this.lblHeaderRows.TabIndex = 3;
            this.lblHeaderRows.Text = "固定前 N 行（头部）：";
            // 
            // numHeaderRows
            // 
            this.numHeaderRows.Location = new System.Drawing.Point(606, 12);
            this.numHeaderRows.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numHeaderRows.Name = "numHeaderRows";
            this.numHeaderRows.Size = new System.Drawing.Size(60, 23);
            this.numHeaderRows.TabIndex = 4;
            this.numHeaderRows.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // lblMaxRows
            // 
            this.lblMaxRows.AutoSize = true;
            this.lblMaxRows.Location = new System.Drawing.Point(460, 48);
            this.lblMaxRows.Name = "lblMaxRows";
            this.lblMaxRows.Size = new System.Drawing.Size(242, 15);
            this.lblMaxRows.TabIndex = 5;
            this.lblMaxRows.Text = "每个拆分文件最大数据行（不含头部）：";
            // 
            // numMaxRows
            // 
            this.numMaxRows.Location = new System.Drawing.Point(708, 44);
            this.numMaxRows.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numMaxRows.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numMaxRows.Name = "numMaxRows";
            this.numMaxRows.Size = new System.Drawing.Size(80, 23);
            this.numMaxRows.TabIndex = 6;
            this.numMaxRows.Value = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            // 
            // lstFiles
            // 
            this.lstFiles.FormattingEnabled = true;
            this.lstFiles.ItemHeight = 15;
            this.lstFiles.Location = new System.Drawing.Point(12, 50);
            this.lstFiles.Name = "lstFiles";
            this.lstFiles.Size = new System.Drawing.Size(440, 379);
            this.lstFiles.TabIndex = 7;
            // 
            // lstLog
            // 
            this.lstLog.FormattingEnabled = true;
            this.lstLog.ItemHeight = 15;
            this.lstLog.Location = new System.Drawing.Point(460, 80);
            this.lstLog.Name = "lstLog";
            this.lstLog.Size = new System.Drawing.Size(328, 349);
            this.lstLog.TabIndex = 8;
            // 
            // openFileDialog
            // 
            this.openFileDialog.Multiselect = true;
            this.openFileDialog.Filter = "Excel 文件|*.xlsx;*.xlsm;*.xls";
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lstLog);
            this.Controls.Add(this.lstFiles);
            this.Controls.Add(this.numMaxRows);
            this.Controls.Add(this.lblMaxRows);
            this.Controls.Add(this.numHeaderRows);
            this.Controls.Add(this.lblHeaderRows);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnSelectOutput);
            this.Controls.Add(this.btnSelectFiles);
            this.Name = "Form1";
            this.Text = "Excel 拆分器";
            ((System.ComponentModel.ISupportInitialize)(this.numMaxRows)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHeaderRows)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
