using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace OfficeUtils
{
    public class SimpleMergeForm : Form
    {
        private Button btnSelectFiles;
        private Button btnSelectOutput;
        private Button btnStart;
        private TextBox txtOutputDir;
        private ListBox lstFiles;
        private RichTextBox rtbLog;
        private SplitContainer MainSplitContainer;
        private SplitContainer BottomSplitContainer;
        private OpenFileDialog openFileDialog;
        private FolderBrowserDialog folderBrowserDialog;

        public SimpleMergeForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            btnSelectFiles = new Button();
            btnSelectOutput = new Button();
            btnStart = new Button();
            txtOutputDir = new TextBox();
            lstFiles = new ListBox();
            rtbLog = new RichTextBox();
            openFileDialog = new OpenFileDialog();
            folderBrowserDialog = new FolderBrowserDialog();
            MainSplitContainer = new SplitContainer();
            BottomSplitContainer = new SplitContainer();
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
            btnSelectFiles.Location = new Point(12, 12);
            btnSelectFiles.Name = "btnSelectFiles";
            btnSelectFiles.Size = new Size(120, 28);
            btnSelectFiles.TabIndex = 0;
            btnSelectFiles.Text = "选择文件...";
            btnSelectFiles.Click += BtnSelectFiles_Click;
            // 
            // btnSelectOutput
            // 
            btnSelectOutput.Location = new Point(140, 12);
            btnSelectOutput.Name = "btnSelectOutput";
            btnSelectOutput.Size = new Size(140, 28);
            btnSelectOutput.TabIndex = 1;
            btnSelectOutput.Text = "选择输出目录...";
            btnSelectOutput.Click += BtnSelectOutput_Click;
            // 
            // btnStart
            // 
            btnStart.Location = new Point(300, 12);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(75, 23);
            btnStart.TabIndex = 2;
            btnStart.Text = "开始简单合并";
            btnStart.Click += BtnStart_Click;
            // 
            // txtOutputDir
            // 
            txtOutputDir.Location = new Point(12, 48);
            txtOutputDir.Name = "txtOutputDir";
            txtOutputDir.ReadOnly = true;
            txtOutputDir.Size = new Size(760, 23);
            txtOutputDir.TabIndex = 3;
            // 
            // lstFiles
            // 
            lstFiles.Dock = DockStyle.Fill;
            lstFiles.Location = new Point(0, 0);
            lstFiles.Name = "lstFiles";
            lstFiles.SelectionMode = SelectionMode.MultiExtended;
            lstFiles.Size = new Size(270, 375);
            lstFiles.TabIndex = 0;
            // 
            // rtbLog
            // 
            rtbLog.Dock = DockStyle.Fill;
            rtbLog.Location = new Point(0, 0);
            rtbLog.Name = "rtbLog";
            rtbLog.ReadOnly = true;
            rtbLog.Size = new Size(510, 375);
            rtbLog.TabIndex = 0;
            rtbLog.Text = "";
            // 
            // openFileDialog
            // 
            openFileDialog.Filter = "Excel 文件|*.xlsx;*.xlsm;*.xls|所有文件|*.*";
            openFileDialog.Multiselect = true;
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
            MainSplitContainer.Panel1.Controls.Add(btnSelectOutput);
            MainSplitContainer.Panel1.Controls.Add(btnStart);
            MainSplitContainer.Panel1.Controls.Add(txtOutputDir);
            // 
            // MainSplitContainer.Panel2
            // 
            MainSplitContainer.Panel2.Controls.Add(BottomSplitContainer);
            MainSplitContainer.Size = new Size(784, 461);
            MainSplitContainer.SplitterDistance = 82;
            MainSplitContainer.TabIndex = 0;
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
            BottomSplitContainer.Panel2.Controls.Add(rtbLog);
            BottomSplitContainer.Size = new Size(784, 375);
            BottomSplitContainer.SplitterDistance = 364;
            BottomSplitContainer.TabIndex = 0;
            // 
            // SimpleMergeForm
            // 
            ClientSize = new Size(784, 461);
            Controls.Add(MainSplitContainer);
            Name = "SimpleMergeForm";
            Text = "Excel 简单合并";
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

        private void BtnSelectFiles_Click(object? sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                lstFiles.Items.Clear();
                foreach (var f in openFileDialog.FileNames)
                    lstFiles.Items.Add(f);
                Log($"已选择 {openFileDialog.FileNames.Length} 个文件");
            }
        }

        private void BtnSelectOutput_Click(object? sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                txtOutputDir.Text = folderBrowserDialog.SelectedPath;
                Log($"输出目录：{txtOutputDir.Text}");
            }
        }

        private async void BtnStart_Click(object? sender, EventArgs e)
        {
            if (lstFiles.Items.Count == 0)
            {
                MessageBox.Show("请先选择文件。", Text);
                return;
            }
            if (string.IsNullOrEmpty(txtOutputDir.Text))
            {
                MessageBox.Show("请先选择输出目录。", Text);
                return;
            }

            btnStart.Enabled = false;
            btnSelectFiles.Enabled = false;
            btnSelectOutput.Enabled = false;

            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var resultWb = new XLWorkbook();
                    foreach (var item in lstFiles.Items.Cast<string>())
                    {
                        var file = item;
                        try
                        {
                            using var wb = new XLWorkbook(file);
                            foreach (var ws in wb.Worksheets)
                            {
                                var name = ws.Name;
                                var targetName = name;
                                int idx = 1;
                                while (resultWb.Worksheets.Any(s => s.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    idx++;
                                    targetName = name + "(" + idx + ")";
                                }
                                var newWs = resultWb.Worksheets.Add(targetName);
                                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                                var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
                                for (int r = 1; r <= lastRow; r++)
                                {
                                    for (int c = 1; c <= lastCol; c++)
                                    {
                                        var src = ws.Cell(r, c);
                                        var dst = newWs.Cell(r, c);
                                        dst.Value = src.Value;
                                        dst.Style = src.Style;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Invoke(new Action(() => Log($"错误：{Path.GetFileName(file)} -> {ex.Message}", error: true)));
                        }
                    }

                    var outPath = Path.Combine(txtOutputDir.Text, "merged_sheets.xlsx");
                    resultWb.SaveAs(outPath);
                    Invoke(new Action(() => Log($"已生成合并文件：{Path.GetFileName(outPath)}")));
                }
                catch (Exception ex)
                {
                    Invoke(new Action(() => Log($"合并失败：{ex.Message}", error: true)));
                }
            });

            Log("简单合并完成。");

            btnStart.Enabled = true;
            btnSelectFiles.Enabled = true;
            btnSelectOutput.Enabled = true;
        }

        private void Log(string message, bool warning = false, bool error = false)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Log(message, warning, error)));
                return;
            }
            var line = $"{DateTime.Now:HH:mm:ss} {message}";
            rtbLog.AppendText(line + Environment.NewLine);
            rtbLog.SelectionStart = rtbLog.Text.Length;
            rtbLog.ScrollToCaret();
            try
            {
                if (error)
                    AppLogger.Logger?.Error(line);
                else if (warning)
                    AppLogger.Logger?.Warning(line);
                else
                    AppLogger.Logger?.Information(line);
            }
            catch { }
        }
    }
}
