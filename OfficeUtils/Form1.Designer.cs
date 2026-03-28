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
        private System.Windows.Forms.RichTextBox lstLog;
        private System.Windows.Forms.NumericUpDown numMaxRows;
        private System.Windows.Forms.NumericUpDown numHeaderRows;
        private System.Windows.Forms.Label lblMaxRows;
        private System.Windows.Forms.Label lblHeaderRowsLabel;
        private System.Windows.Forms.Label lblOutputDirLabel;
        private System.Windows.Forms.TextBox txtOutputDir;
        private System.Windows.Forms.Label lblGroupCols;
        private System.Windows.Forms.TextBox txtGroupCols;
        private System.Windows.Forms.CheckBox chkSplitToFiles;
        private System.Windows.Forms.ContextMenuStrip ctxListCopy;
        private System.Windows.Forms.ToolStripMenuItem menuCopy;
        private System.Windows.Forms.ToolStripMenuItem menuCopyAll;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SplitContainer MainSplitContainer;
        private System.Windows.Forms.SplitContainer BottomSplitContainer;

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            btnSelectFiles = new Button();
            btnSelectOutput = new Button();
            btnStart = new Button();
            lstFiles = new ListBox();
            ctxListCopy = new ContextMenuStrip(components);
            menuCopy = new ToolStripMenuItem();
            menuCopyAll = new ToolStripMenuItem();
            lstLog = new RichTextBox();
            numMaxRows = new NumericUpDown();
            numHeaderRows = new NumericUpDown();
            lblMaxRows = new Label();
            lblOutputDirLabel = new Label();
            txtOutputDir = new TextBox();
            lblGroupCols = new Label();
            txtGroupCols = new TextBox();
            lblHeaderRowsLabel = new Label();
            folderBrowserDialog = new FolderBrowserDialog();
            openFileDialog = new OpenFileDialog();
            chkSplitToFiles = new CheckBox();
            BottomSplitContainer = new SplitContainer();
            MainSplitContainer = new SplitContainer();
            ctxListCopy.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)numMaxRows).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numHeaderRows).BeginInit();
            ((System.ComponentModel.ISupportInitialize)BottomSplitContainer).BeginInit();
            BottomSplitContainer.Panel1.SuspendLayout();
            BottomSplitContainer.Panel2.SuspendLayout();
            BottomSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)MainSplitContainer).BeginInit();
            MainSplitContainer.Panel1.SuspendLayout();
            MainSplitContainer.Panel2.SuspendLayout();
            MainSplitContainer.SuspendLayout();
            SuspendLayout();
            // 
            // btnSelectFiles
            // 
            btnSelectFiles.Location = new Point(12, 12);
            btnSelectFiles.Name = "btnSelectFiles";
            btnSelectFiles.Size = new Size(150, 27);
            btnSelectFiles.TabIndex = 0;
            btnSelectFiles.Text = "选择 Excel 文件...";
            btnSelectFiles.UseVisualStyleBackColor = true;
            btnSelectFiles.Click += btnSelectFiles_Click;
            // 
            // btnSelectOutput
            // 
            btnSelectOutput.Location = new Point(168, 12);
            btnSelectOutput.Name = "btnSelectOutput";
            btnSelectOutput.Size = new Size(150, 27);
            btnSelectOutput.TabIndex = 1;
            btnSelectOutput.Text = "选择输出目录...";
            btnSelectOutput.UseVisualStyleBackColor = true;
            btnSelectOutput.Click += btnSelectOutput_Click;
            // 
            // btnStart
            // 
            btnStart.Location = new Point(324, 12);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(120, 27);
            btnStart.TabIndex = 2;
            btnStart.Text = "确定拆分";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // lstFiles
            // 
            lstFiles.ContextMenuStrip = ctxListCopy;
            lstFiles.Dock = DockStyle.Fill;
            lstFiles.FormattingEnabled = true;
            lstFiles.Location = new Point(0, 0);
            lstFiles.Name = "lstFiles";
            lstFiles.SelectionMode = SelectionMode.MultiExtended;
            lstFiles.Size = new Size(364, 451);
            lstFiles.TabIndex = 7;
            lstFiles.KeyDown += lstList_KeyDown;
            // 
            // ctxListCopy
            // 
            ctxListCopy.Items.AddRange(new ToolStripItem[] { menuCopy, menuCopyAll });
            ctxListCopy.Name = "ctxListCopy";
            ctxListCopy.Size = new Size(141, 48);
            // 
            // menuCopy
            // 
            menuCopy.Name = "menuCopy";
            menuCopy.Size = new Size(140, 22);
            menuCopy.Text = "复制所选(&C)";
            menuCopy.Click += menuCopy_Click;
            // 
            // menuCopyAll
            // 
            menuCopyAll.Name = "menuCopyAll";
            menuCopyAll.Size = new Size(140, 22);
            menuCopyAll.Text = "全部复制(&A)";
            menuCopyAll.Click += menuCopyAll_Click;
            // 
            // lstLog
            // 
            lstLog.BackColor = Color.White;
            lstLog.ContextMenuStrip = ctxListCopy;
            lstLog.Dock = DockStyle.Fill;
            lstLog.Font = new Font("Segoe UI", 9F);
            lstLog.Location = new Point(0, 0);
            lstLog.Name = "lstLog";
            lstLog.ReadOnly = true;
            lstLog.Size = new Size(733, 451);
            lstLog.TabIndex = 8;
            lstLog.Text = "";
            lstLog.KeyDown += lstList_KeyDown;
            // 
            // numMaxRows
            // 
            numMaxRows.Location = new Point(260, 108);
            numMaxRows.Maximum = new decimal(new int[] { 1000000, 0, 0, 0 });
            numMaxRows.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numMaxRows.Name = "numMaxRows";
            numMaxRows.Size = new Size(80, 23);
            numMaxRows.TabIndex = 8;
            numMaxRows.Value = new decimal(new int[] { 1998, 0, 0, 0 });
            // 
            // numHeaderRows
            // 
            numHeaderRows.Location = new Point(158, 80);
            numHeaderRows.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            numHeaderRows.Name = "numHeaderRows";
            numHeaderRows.Size = new Size(60, 23);
            numHeaderRows.TabIndex = 6;
            numHeaderRows.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // lblMaxRows
            // 
            lblMaxRows.AutoSize = true;
            lblMaxRows.Location = new Point(12, 112);
            lblMaxRows.Name = "lblMaxRows";
            lblMaxRows.Size = new Size(224, 17);
            lblMaxRows.TabIndex = 7;
            lblMaxRows.Text = "每个拆分文件最大数据行（不含头部）：";
            // 
            // lblOutputDirLabel
            // 
            lblOutputDirLabel.AutoSize = true;
            lblOutputDirLabel.Location = new Point(12, 52);
            lblOutputDirLabel.Name = "lblOutputDirLabel";
            lblOutputDirLabel.Size = new Size(68, 17);
            lblOutputDirLabel.TabIndex = 3;
            lblOutputDirLabel.Text = "输出目录：";
            // 
            // txtOutputDir
            // 
            txtOutputDir.Location = new Point(90, 48);
            txtOutputDir.Name = "txtOutputDir";
            txtOutputDir.ReadOnly = true;
            txtOutputDir.Size = new Size(698, 23);
            txtOutputDir.TabIndex = 4;
            // 
            // lblGroupCols
            // 
            lblGroupCols.AutoSize = true;
            lblGroupCols.Location = new Point(360, 84);
            lblGroupCols.Name = "lblGroupCols";
            lblGroupCols.Size = new Size(140, 17);
            lblGroupCols.TabIndex = 9;
            lblGroupCols.Text = "分组列号（逗号分隔）：";
            // 
            // txtGroupCols
            // 
            txtGroupCols.Location = new Point(520, 80);
            txtGroupCols.Name = "txtGroupCols";
            txtGroupCols.Size = new Size(268, 23);
            txtGroupCols.TabIndex = 10;
            txtGroupCols.Text = "A,B";
            txtGroupCols.Leave += txtGroupCols_Leave;
            // 
            // lblHeaderRowsLabel
            // 
            lblHeaderRowsLabel.AutoSize = true;
            lblHeaderRowsLabel.Location = new Point(12, 84);
            lblHeaderRowsLabel.Name = "lblHeaderRowsLabel";
            lblHeaderRowsLabel.Size = new Size(134, 17);
            lblHeaderRowsLabel.TabIndex = 5;
            lblHeaderRowsLabel.Text = "固定前 N 行（头部）：";
            // 
            // openFileDialog
            // 
            openFileDialog.Filter = "Excel 文件|*.xlsx;*.xlsm;*.xls";
            openFileDialog.Multiselect = true;
            // 
            // chkSplitToFiles
            // 
            chkSplitToFiles.FlatStyle = FlatStyle.Flat;
            chkSplitToFiles.ForeColor = Color.DimGray;
            chkSplitToFiles.Location = new Point(469, 16);
            chkSplitToFiles.Name = "chkSplitToFiles";
            chkSplitToFiles.Size = new Size(240, 21);
            chkSplitToFiles.TabIndex = 11;
            chkSplitToFiles.Text = "拆分成多个文件（每部分为单独文件）";
            chkSplitToFiles.CheckedChanged += chkSplitToFiles_CheckedChanged;
            // 
            // BottomSplitContainer
            // 
            BottomSplitContainer.Dock = DockStyle.Fill;
            BottomSplitContainer.Location = new Point(0, 0);
            BottomSplitContainer.Name = "BottomSplitContainer";
            // 
            // BottomSplitContainer.Panel1
            // 
            BottomSplitContainer.Panel1.Controls.Add(lstFiles);
            // 
            // BottomSplitContainer.Panel2
            // 
            BottomSplitContainer.Panel2.Controls.Add(lstLog);
            BottomSplitContainer.Panel2.Paint += splitContainer1_Panel2_Paint;
            BottomSplitContainer.Size = new Size(1102, 451);
            BottomSplitContainer.SplitterDistance = 364;
            BottomSplitContainer.SplitterWidth = 5;
            BottomSplitContainer.TabIndex = 12;
            BottomSplitContainer.TabStop = false;
            // 
            // MainSplitContainer
            // 
            MainSplitContainer.Dock = DockStyle.Fill;
            MainSplitContainer.FixedPanel = FixedPanel.Panel1;
            MainSplitContainer.Location = new Point(0, 0);
            MainSplitContainer.MinimumSize = new Size(800, 600);
            MainSplitContainer.Name = "MainSplitContainer";
            MainSplitContainer.Orientation = Orientation.Horizontal;
            // 
            // MainSplitContainer.Panel1
            // 
            MainSplitContainer.Panel1.Controls.Add(btnSelectFiles);
            MainSplitContainer.Panel1.Controls.Add(btnSelectOutput);
            MainSplitContainer.Panel1.Controls.Add(btnStart);
            MainSplitContainer.Panel1.Controls.Add(lblOutputDirLabel);
            MainSplitContainer.Panel1.Controls.Add(txtOutputDir);
            MainSplitContainer.Panel1.Controls.Add(lblGroupCols);
            MainSplitContainer.Panel1.Controls.Add(txtGroupCols);
            MainSplitContainer.Panel1.Controls.Add(chkSplitToFiles);
            MainSplitContainer.Panel1.Controls.Add(lblHeaderRowsLabel);
            MainSplitContainer.Panel1.Controls.Add(numHeaderRows);
            MainSplitContainer.Panel1.Controls.Add(lblMaxRows);
            MainSplitContainer.Panel1.Controls.Add(numMaxRows);
            // 
            // MainSplitContainer.Panel2
            // 
            MainSplitContainer.Panel2.Controls.Add(BottomSplitContainer);
            MainSplitContainer.Size = new Size(1102, 600);
            MainSplitContainer.SplitterDistance = 145;
            MainSplitContainer.TabIndex = 13;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1102, 520);
            Controls.Add(MainSplitContainer);
            Name = "Form1";
            Text = "Excel 拆分器";
            ctxListCopy.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)numMaxRows).EndInit();
            ((System.ComponentModel.ISupportInitialize)numHeaderRows).EndInit();
            BottomSplitContainer.Panel1.ResumeLayout(false);
            BottomSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)BottomSplitContainer).EndInit();
            BottomSplitContainer.ResumeLayout(false);
            MainSplitContainer.Panel1.ResumeLayout(false);
            MainSplitContainer.Panel1.PerformLayout();
            MainSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)MainSplitContainer).EndInit();
            MainSplitContainer.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion


    }
}
