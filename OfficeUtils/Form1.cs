using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace OfficeUtils
{
    public partial class Form1 : Form
    {
        private string outputDirectory;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelectFiles_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                lstFiles.Items.Clear();
                foreach (var f in openFileDialog.FileNames)
                    lstFiles.Items.Add(f);
                Log($"已选择 {openFileDialog.FileNames.Length} 个文件");
            }
        }

        private void btnSelectOutput_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                outputDirectory = folderBrowserDialog.SelectedPath;
                Log($"输出目录：{outputDirectory}");
            }
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            if (lstFiles.Items.Count == 0)
            {
                MessageBox.Show("请先选择要拆分的 Excel 文件。", Text);
                return;
            }
            if (string.IsNullOrEmpty(outputDirectory))
            {
                MessageBox.Show("请先选择输出目录。", Text);
                return;
            }

            int headerRows = (int)numHeaderRows.Value;
            int maxRowsPerFile = (int)numMaxRows.Value;
            var files = lstFiles.Items.Cast<string>().ToArray();

            btnStart.Enabled = false;
            btnSelectFiles.Enabled = false;
            btnSelectOutput.Enabled = false;
            try
            {
                await Task.Run(() =>
                {
                    foreach (var file in files)
                    {
                        try
                        {
                            SplitFile(file, maxRowsPerFile, headerRows, outputDirectory);
                            Invoke(new Action(() => Log($"拆分完成：{Path.GetFileName(file)}")));
                        }
                        catch (Exception ex)
                        {
                            Invoke(new Action(() => Log($"错误：{Path.GetFileName(file)} -> {ex.Message}")));
                        }
                    }
                });

                MessageBox.Show("全部拆分任务已完成。", Text);
            }
            finally
            {
                btnStart.Enabled = true;
                btnSelectFiles.Enabled = true;
                btnSelectOutput.Enabled = true;
            }
        }

        private void SplitFile(string inputPath, int maxRowsPerFile, int headerRows, string outputDir)
        {
            Invoke(new Action(() => Log($"开始拆分（为 Sheet 生成分片）：{Path.GetFileName(inputPath)}")));
            using var wb = new XLWorkbook(inputPath);
            var ws = wb.Worksheets.First(); // 仅处理第一个 sheet
            var lastRowUsed = ws.LastRowUsed()?.RowNumber() ?? 0;
            var lastColUsed = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

            if (lastRowUsed == 0 || lastColUsed == 0)
            {
                Invoke(new Action(() => Log($"跳过空表：{Path.GetFileName(inputPath)}")));
                return;
            }

            int dataStart = headerRows + 1;
            int remainder = Math.Max(0, lastRowUsed - headerRows);

            // 如果没有需要拆分的数据，则直接复制为输出文件
            if (remainder == 0 || remainder <= maxRowsPerFile)
            {
                var destName = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(inputPath)}_sheets{Path.GetExtension(inputPath)}");
                wb.SaveAs(destName);
                Invoke(new Action(() => Log($"无需拆分或数据量不足，已复制为：{Path.GetFileName(destName)}")));
                return;
            }

            int partIndex = 1;

            double[] columnWidths = new double[lastColUsed + 1];
            for (int c = 1; c <= lastColUsed; c++)
                columnWidths[c] = ws.Column(c).Width;

            for (int offset = 0; offset < remainder; offset += maxRowsPerFile)
            {
                int chunkStart = dataStart + offset;
                int chunkEnd = Math.Min(chunkStart + maxRowsPerFile - 1, lastRowUsed);

                // 生成唯一的 sheet 名称
                string baseName = ws.Name + "_part" + partIndex;
                string newSheetName = baseName;
                int tryIndex = 1;
                while (wb.Worksheets.Any(s => s.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    tryIndex++;
                    newSheetName = baseName + "(" + tryIndex + ")";
                }

                var newWs = wb.Worksheets.Add(newSheetName);

                // 复制列宽
                for (int c = 1; c <= lastColUsed; c++)
                    newWs.Column(c).Width = columnWidths[c];

                int destRow = 1;

                // 复制 headerRows
                for (int r = 1; r <= headerRows; r++)
                {
                    for (int c = 1; c <= lastColUsed; c++)
                    {
                        var srcCell = ws.Cell(r, c);
                        var dstCell = newWs.Cell(destRow, c);
                        dstCell.Value = srcCell.Value;
                        dstCell.Style = srcCell.Style;
                    }
                    destRow++;
                }

                // 复制 chunk 行
                for (int r = chunkStart; r <= chunkEnd; r++)
                {
                    for (int c = 1; c <= lastColUsed; c++)
                    {
                        var srcCell = ws.Cell(r, c);
                        var dstCell = newWs.Cell(destRow, c);
                        dstCell.Value = srcCell.Value;
                        dstCell.Style = srcCell.Style;
                    }
                    destRow++;
                }

                Invoke(new Action(() => Log($"已生成 sheet：{newSheetName} (行 {chunkStart}-{chunkEnd}，含头部 {headerRows} 行)")));
                partIndex++;
            }

            // 保存为一个新文件，保留原文件不变
            string outFile = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(inputPath)}_sheets{Path.GetExtension(inputPath)}");
            wb.SaveAs(outFile);
            Invoke(new Action(() => Log($"已保存带分片的文件：{Path.GetFileName(outFile)}")));
        }

        private void Log(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Log(message)));
                return;
            }
            lstLog.Items.Add($"{DateTime.Now:HH:mm:ss} {message}");
            lstLog.TopIndex = lstLog.Items.Count - 1;
        }
    }
}
