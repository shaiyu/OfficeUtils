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

        private void InitializeComponent()
        {
            this.btnSelectFiles = new Button();
            this.btnStartUpload = new Button();
            this.txtUrl = new TextBox();
            this.lstFiles = new ListBox();
            this.lstResults = new ListBox();
            this.openFileDialog = new OpenFileDialog();

            // btnSelectFiles
            this.btnSelectFiles.Text = "选择文件...";
            this.btnSelectFiles.Location = new System.Drawing.Point(8, 8);
            this.btnSelectFiles.Size = new System.Drawing.Size(100, 28);
            this.btnSelectFiles.Click += BtnSelectFiles_Click;

            // txtUrl
            this.txtUrl.Location = new System.Drawing.Point(120, 12);
            this.txtUrl.Size = new System.Drawing.Size(640, 23);
            this.txtUrl.Text = DefaultUrl;

            // lblCookie
            this.lblCookie = new Label();
            this.lblCookie.Location = new System.Drawing.Point(8, 44);
            this.lblCookie.Size = new System.Drawing.Size(100, 23);
            this.lblCookie.Text = "Cookie:";

            // txtCookie
            this.txtCookie = new TextBox();
            this.txtCookie.Location = new System.Drawing.Point(120, 44);
            this.txtCookie.Size = new System.Drawing.Size(640, 23);
            this.txtCookie.Text = string.Empty;

            // btnStartUpload
            this.btnStartUpload.Text = "开始上传";
            this.btnStartUpload.Location = new System.Drawing.Point(770, 8);
            this.btnStartUpload.Size = new System.Drawing.Size(100, 28);
            this.btnStartUpload.Click += BtnStartUpload_Click;

            // lstFiles
            this.lstFiles.Location = new System.Drawing.Point(8, 80);
            this.lstFiles.Size = new System.Drawing.Size(420, 420);
            this.lstFiles.SelectionMode = SelectionMode.MultiExtended;

            // lstResults
            this.lstResults.Location = new System.Drawing.Point(440, 80);
            this.lstResults.Size = new System.Drawing.Size(430, 420);

            // openFileDialog
            this.openFileDialog.Multiselect = true;
            this.openFileDialog.Filter = "Excel 文件|*.xlsx;*.xlsm;*.xls|所有文件|*.*";

            // UploadForm
            this.Controls.Add(this.btnSelectFiles);
            this.Controls.Add(this.txtUrl);
            this.Controls.Add(this.lblCookie);
            this.Controls.Add(this.txtCookie);
            this.Controls.Add(this.btnStartUpload);
            this.Controls.Add(this.lstFiles);
            this.Controls.Add(this.lstResults);
            this.Text = "批量上传";
            this.ClientSize = new System.Drawing.Size(880, 500);
        }
    }
}
