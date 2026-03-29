using System;
using System.Windows.Forms;

namespace OfficeUtils
{
    partial class UploadForm
    {
        private Button btnSelectFiles;
        private Button btnStartUpload;
        private TextBox txtUrl;
        private Label lblCookie;
        private TextBox txtCookie;
        private ListBox lstFiles;
        private ListBox lstResults;
        private OpenFileDialog openFileDialog;
        private System.Windows.Forms.SplitContainer MainSplitContainer;
        private System.Windows.Forms.SplitContainer BottomSplitContainer;

        private void InitializeComponent()
        {
            btnSelectFiles = new Button();
            btnStartUpload = new Button();
            txtUrl = new TextBox();
            lstFiles = new ListBox();
            lstResults = new ListBox();
            MainSplitContainer = new SplitContainer();
            lblCookie = new Label();
            txtCookie = new TextBox();
            BottomSplitContainer = new SplitContainer();
            openFileDialog = new OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)MainSplitContainer).BeginInit();
            MainSplitContainer.Panel1.SuspendLayout();
            MainSplitContainer.Panel2.SuspendLayout();
            MainSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)BottomSplitContainer).BeginInit();
            BottomSplitContainer.Panel1.SuspendLayout();
            BottomSplitContainer.Panel2.SuspendLayout();
            BottomSplitContainer.SuspendLayout();
            SuspendLayout();
            // 
            // btnSelectFiles
            // 
            btnSelectFiles.Location = new Point(8, 8);
            btnSelectFiles.Name = "btnSelectFiles";
            btnSelectFiles.Size = new Size(100, 28);
            btnSelectFiles.TabIndex = 0;
            btnSelectFiles.Text = "选择文件...";
            btnSelectFiles.Click += BtnSelectFiles_Click;
            // 
            // btnStartUpload
            // 
            btnStartUpload.Location = new Point(770, 8);
            btnStartUpload.Name = "btnStartUpload";
            btnStartUpload.Size = new Size(100, 28);
            btnStartUpload.TabIndex = 4;
            btnStartUpload.Text = "开始上传";
            btnStartUpload.Click += BtnStartUpload_Click;
            // 
            // txtUrl
            // 
            txtUrl.Location = new Point(120, 12);
            txtUrl.Name = "txtUrl";
            txtUrl.Size = new Size(640, 23);
            txtUrl.TabIndex = 1;
            // 
            // lstFiles
            // 
            lstFiles.Dock = DockStyle.Fill;
            lstFiles.Location = new Point(0, 0);
            lstFiles.Name = "lstFiles";
            lstFiles.SelectionMode = SelectionMode.MultiExtended;
            lstFiles.Size = new Size(395, 414);
            lstFiles.TabIndex = 0;
            lstFiles.SelectedIndexChanged += lstFiles_SelectedIndexChanged;
            // 
            // lstResults
            // 
            lstResults.Dock = DockStyle.Fill;
            lstResults.Location = new Point(0, 0);
            lstResults.Name = "lstResults";
            lstResults.Size = new Size(605, 414);
            lstResults.TabIndex = 0;
            lstResults.SelectedIndexChanged += lstResults_SelectedIndexChanged;
            // 
            // MainSplitContainer
            // 
            MainSplitContainer.Dock = DockStyle.Fill;
            MainSplitContainer.FixedPanel = FixedPanel.Panel1;
            MainSplitContainer.Location = new Point(0, 0);
            MainSplitContainer.Name = "MainSplitContainer";
            MainSplitContainer.Orientation = Orientation.Horizontal;
            // 
            // MainSplitContainer.Panel1
            // 
            MainSplitContainer.Panel1.Controls.Add(btnSelectFiles);
            MainSplitContainer.Panel1.Controls.Add(txtUrl);
            MainSplitContainer.Panel1.Controls.Add(lblCookie);
            MainSplitContainer.Panel1.Controls.Add(txtCookie);
            MainSplitContainer.Panel1.Controls.Add(btnStartUpload);
            // 
            // MainSplitContainer.Panel2
            // 
            MainSplitContainer.Panel2.Controls.Add(BottomSplitContainer);
            MainSplitContainer.Panel2.Paint += MainSplitContainer_Panel2_Paint;
            MainSplitContainer.Size = new Size(1005, 511);
            MainSplitContainer.SplitterDistance = 93;
            MainSplitContainer.TabIndex = 0;
            // 
            // lblCookie
            // 
            lblCookie.Location = new Point(8, 44);
            lblCookie.Name = "lblCookie";
            lblCookie.Size = new Size(100, 23);
            lblCookie.TabIndex = 2;
            lblCookie.Text = "Cookie:";
            // 
            // txtCookie
            // 
            txtCookie.Location = new Point(120, 44);
            txtCookie.Name = "txtCookie";
            txtCookie.Size = new Size(640, 23);
            txtCookie.TabIndex = 3;
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
            BottomSplitContainer.Panel2.Controls.Add(lstResults);
            BottomSplitContainer.Size = new Size(1005, 414);
            BottomSplitContainer.SplitterDistance = 364;
            BottomSplitContainer.SplitterWidth = 5;
            BottomSplitContainer.TabIndex = 0;
            // 
            // openFileDialog
            // 
            openFileDialog.Filter = "Excel 文件|*.xlsx;*.xlsm;*.xls|所有文件|*.*";
            openFileDialog.Multiselect = true;
            // 
            // UploadForm
            // 
            BackColor = Color.WhiteSmoke;
            ClientSize = new Size(1005, 511);
            Controls.Add(MainSplitContainer);
            Font = new Font("Segoe UI", 9F);
            Name = "UploadForm";
            Text = "批量上传";
            MainSplitContainer.Panel1.ResumeLayout(false);
            MainSplitContainer.Panel1.PerformLayout();
            MainSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)MainSplitContainer).EndInit();
            MainSplitContainer.ResumeLayout(false);
            BottomSplitContainer.Panel1.ResumeLayout(false);
            BottomSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)BottomSplitContainer).EndInit();
            BottomSplitContainer.ResumeLayout(false);
            ResumeLayout(false);
        }
    }
}
