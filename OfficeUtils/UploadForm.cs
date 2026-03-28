using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeUtils
{
    public class UploadForm : Form
    {
        private Button btnSelectFiles;
        private Button btnStartUpload;
        private TextBox txtUrl;
        private ListBox lstFiles;
        private ListBox lstResults;
        private OpenFileDialog openFileDialog;

        private const string DefaultUrl = "https://jproapi.ningmengyun.com/api/Voucher/Import?autoAddAsub=false&isAutoVnumAtRepeat=false&appasid=200805416I81l81";

        public UploadForm()
        {
            InitializeComponent();
        }

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

            // btnStartUpload
            this.btnStartUpload.Text = "开始上传";
            this.btnStartUpload.Location = new System.Drawing.Point(770, 8);
            this.btnStartUpload.Size = new System.Drawing.Size(100, 28);
            this.btnStartUpload.Click += BtnStartUpload_Click;

            // lstFiles
            this.lstFiles.Location = new System.Drawing.Point(8, 48);
            this.lstFiles.Size = new System.Drawing.Size(420, 440);
            this.lstFiles.SelectionMode = SelectionMode.MultiExtended;

            // lstResults
            this.lstResults.Location = new System.Drawing.Point(440, 48);
            this.lstResults.Size = new System.Drawing.Size(430, 440);

            // openFileDialog
            this.openFileDialog.Multiselect = true;
            this.openFileDialog.Filter = "Excel 文件|*.xlsx;*.xlsm;*.xls|所有文件|*.*";

            // UploadForm
            this.Controls.Add(this.btnSelectFiles);
            this.Controls.Add(this.txtUrl);
            this.Controls.Add(this.btnStartUpload);
            this.Controls.Add(this.lstFiles);
            this.Controls.Add(this.lstResults);
            this.Text = "批量上传";
            this.ClientSize = new System.Drawing.Size(880, 500);
        }

        private void BtnSelectFiles_Click(object? sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                lstFiles.Items.Clear();
                foreach (var f in openFileDialog.FileNames)
                    lstFiles.Items.Add(f);
                lstResults.Items.Add($"已选择 {openFileDialog.FileNames.Length} 个文件");
            }
        }

        private async void BtnStartUpload_Click(object? sender, EventArgs e)
        {
            if (lstFiles.Items.Count == 0)
            {
                MessageBox.Show("请先选择要上传的文件。", Text);
                return;
            }
            var url = txtUrl.Text?.Trim();
            if (string.IsNullOrEmpty(url))
            {
                MessageBox.Show("请输入接口地址。", Text);
                return;
            }

            btnStartUpload.Enabled = false;
            btnSelectFiles.Enabled = false;
            lstResults.Items.Add($"开始上传 {lstFiles.Items.Count} 个文件...");

            using var client = new HttpClient();

            foreach (var item in lstFiles.Items.Cast<string>())
            {
                var filePath = item;
                try
                {
                    lstResults.Items.Add($"上传中: {Path.GetFileName(filePath)}");
                    using var content = new MultipartFormDataContent();
                    using var fs = File.OpenRead(filePath);
                    var streamContent = new StreamContent(fs);
                    // add as form-data with name 'file' and proper filename
                    content.Add(streamContent, "file", Path.GetFileName(filePath));

                    var resp = await client.PostAsync(url, content);
                    var respText = await resp.Content.ReadAsStringAsync();
                    lstResults.Items.Add($"{DateTime.Now:HH:mm:ss} {Path.GetFileName(filePath)} -> {resp.StatusCode} {respText}");
                }
                catch (Exception ex)
                {
                    lstResults.Items.Add($"{DateTime.Now:HH:mm:ss} {Path.GetFileName(filePath)} -> 错误: {ex.Message}");
                }
            }

            lstResults.Items.Add("上传完成。");
            btnStartUpload.Enabled = true;
            btnSelectFiles.Enabled = true;
        }
    }
}

