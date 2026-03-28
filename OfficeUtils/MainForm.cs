using System;
using System.Linq;
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
            tabControl = new TabControl();
            tabExcelSplit = new TabPage();
            tabBatchUpload = new TabPage();
            tabControl.SuspendLayout();
            SuspendLayout();
            // 
            // tabControl
            // 
            tabControl.Controls.Add(tabExcelSplit);
            tabControl.Controls.Add(tabBatchUpload);
            tabControl.Dock = DockStyle.Fill;
            tabControl.Location = new Point(0, 0);
            tabControl.Name = "tabControl";
            tabControl.SelectedIndex = 0;
            tabControl.Size = new Size(900, 600);
            tabControl.TabIndex = 0;
            // 
            // tabExcelSplit
            // 
            tabExcelSplit.Location = new Point(4, 26);
            tabExcelSplit.Name = "tabExcelSplit";
            tabExcelSplit.Padding = new Padding(6);
            tabExcelSplit.Size = new Size(892, 570);
            tabExcelSplit.TabIndex = 0;
            tabExcelSplit.Text = "Excel 拆分";
            // 
            // tabBatchUpload
            // 
            tabBatchUpload.Location = new Point(4, 26);
            tabBatchUpload.Name = "tabBatchUpload";
            tabBatchUpload.Padding = new Padding(6);
            tabBatchUpload.Size = new Size(192, 70);
            tabBatchUpload.TabIndex = 1;
            tabBatchUpload.Text = "批量上传";
            // 
            // MainForm
            // 
            ClientSize = new Size(900, 600);
            Controls.Add(tabControl);
            MinimumSize = new Size(800, 600);
            Name = "MainForm";
            Text = "OfficeUtils";
            Load += MainForm_Load;
            tabControl.ResumeLayout(false);
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
                try
                {
                    var ctrl = form1.Controls.Find("lstLog", true).FirstOrDefault();
                    if (ctrl is RichTextBox rtb)
                    {
                        AppLogger.AttachRichTextBox(rtb);
                    }
                }
                catch
                {
                    // ignore
                }
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
