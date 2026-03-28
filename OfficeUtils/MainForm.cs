using System;
using System.Windows.Forms;

namespace OfficeUtils
{
    public class MainForm : Form
    {
        private TabControl tabControl;
        private TabPage tabExcelSplit;
        private TabPage tabBatchUpload;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.tabControl = new TabControl();
            this.tabExcelSplit = new TabPage();
            this.tabBatchUpload = new TabPage();

            SuspendLayout();

            // tabControl
            this.tabControl.Dock = DockStyle.Fill;
            this.tabControl.Controls.Add(this.tabExcelSplit);
            this.tabControl.Controls.Add(this.tabBatchUpload);

            // tabExcelSplit
            this.tabExcelSplit.Text = "Excel 拆分";
            this.tabExcelSplit.Padding = new Padding(6);

            // tabBatchUpload
            this.tabBatchUpload.Text = "批量上传";
            this.tabBatchUpload.Padding = new Padding(6);

            // MainForm
            this.ClientSize = new System.Drawing.Size(900, 600);
            this.Controls.Add(this.tabControl);
            this.Text = "OfficeUtils";
            this.Load += MainForm_Load;

            ResumeLayout(false);
        }

        private void MainForm_Load(object? sender, EventArgs e)
        {
            // Host existing Form1 in the first tab
            try
            {
                var form1 = new Form1();
                form1.TopLevel = false;
                form1.FormBorderStyle = FormBorderStyle.None;
                form1.Dock = DockStyle.Fill;
                this.tabExcelSplit.Controls.Add(form1);
                form1.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"无法加载 Excel 拆分模块：{ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Host UploadForm (empty) in second tab
            try
            {
                var upload = new UploadForm();
                upload.TopLevel = false;
                upload.FormBorderStyle = FormBorderStyle.None;
                upload.Dock = DockStyle.Fill;
                this.tabBatchUpload.Controls.Add(upload);
                upload.Show();
            }
            catch
            {
                // ignore
            }
        }
    }
}
