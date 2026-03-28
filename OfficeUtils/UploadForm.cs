using Serilog.Core;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeUtils
{
    public partial class UploadForm : Form
    {
        // Designer will provide control fields and InitializeComponent

        private const string DefaultUrl = "https://jproapi.ningmengyun.com/api/Voucher/Import?autoAddAsub=false&isAutoVnumAtRepeat=false&appasid=200805416I81l81";


        public UploadForm()
        {
            InitializeComponent();
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
            var cookieValue = txtCookie?.Text?.Trim();

            AppLogger.Logger?.Information("Start uploading {Count} files to {Url}", lstFiles.Items.Count, url);

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

                    // attach cookie header if provided
                    if (!string.IsNullOrEmpty(cookieValue))
                    {
                        if (client.DefaultRequestHeaders.Contains("Cookie"))
                            client.DefaultRequestHeaders.Remove("Cookie");
                        client.DefaultRequestHeaders.Add("Cookie", cookieValue);
                    }

                    var resp = await client.PostAsync(url, content);
                    var respText = await resp.Content.ReadAsStringAsync();
                    lstResults.Items.Add($"{DateTime.Now:HH:mm:ss} {Path.GetFileName(filePath)} -> {resp.StatusCode} {respText}");
                    var logLine = $"{DateTime.Now:HH:mm:ss} {Path.GetFileName(filePath)} -> {resp.StatusCode} {respText}";
                    lstResults.TopIndex = lstResults.Items.Count - 1;
                    AppLogger.Logger?.Information(logLine);
                }
                catch (Exception ex)
                {
                    var errLine = $"{DateTime.Now:HH:mm:ss} {Path.GetFileName(filePath)} -> 错误: {ex.Message}";
                    lstResults.Items.Add(errLine);
                    lstResults.TopIndex = lstResults.Items.Count - 1;
                    AppLogger.Logger?.Error(ex, errLine);
                }
            }

            lstResults.Items.Add("上传完成。");
            AppLogger.Logger?.Information("Batch upload completed");
            btnStartUpload.Enabled = true;
            btnSelectFiles.Enabled = true;
        }

        private void lstFiles_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void MainSplitContainer_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lstResults_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

