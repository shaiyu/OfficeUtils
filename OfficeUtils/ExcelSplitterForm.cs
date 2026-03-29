using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace OfficeUtils
{
    public partial class ExcelSplitterForm : Form
    {
        private string outputDirectory;

        public ExcelSplitterForm()
        {
            InitializeComponent();
        }

        private void txtGroupCols_Leave(object? sender, EventArgs e)
        {
            if (txtGroupCols == null) return;
            var text = txtGroupCols.Text ?? string.Empty;
            // 替换中文逗号为英文逗号
            text = text.Replace('，', ',');
            // 规范化：去除重复逗号与空格，转换为大写
            var parts = text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim().ToUpperInvariant())
                .ToArray();

            if (parts.Length == 0)
            {
                txtGroupCols.Text = string.Empty;
                return;
            }

            // 验证每个部分为有效的 Excel 列名称（A..XFD）
            foreach (var p in parts)
            {
                if (!IsValidExcelColumnName(p))
                {
                    MessageBox.Show($"无效的列名：{p}。请输入 Excel 列名（例如 A 或 AB），多个列用逗号分隔。", Text);
                    txtGroupCols.Focus();
                    return;
                }
            }

            txtGroupCols.Text = string.Join(',', parts);
        }

        private static bool IsValidExcelColumnName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            // 最大列为 XFD (16384)
            int col = ToColumnNumber(name);
            return col >= 1 && col <= 16384;
        }

        private static int ToColumnNumber(string colName)
        {
            if (string.IsNullOrEmpty(colName)) return 0;
            colName = colName.Trim().ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < colName.Length; i++)
            {
                char c = colName[i];
                if (c < 'A' || c > 'Z') return 0;
                sum = sum * 26 + (c - 'A' + 1);
            }
            return sum;
        }

        private void btnSelectFiles_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                lstFiles.Items.Clear();
                foreach (var f in openFileDialog.FileNames)
                    lstFiles.Items.Add(f);
                // 默认输出目录为第一个选中文件的目录
                var firstDir = Path.GetDirectoryName(openFileDialog.FileNames.First());
                if (!string.IsNullOrEmpty(firstDir))
                {
                    outputDirectory = firstDir;
                    if (txtOutputDir != null)
                        txtOutputDir.Text = outputDirectory;
                }
                Log($"已选择 {openFileDialog.FileNames.Length} 个文件，默认输出目录：{outputDirectory}");
                AppLogger.Logger?.Information("Selected {Count} files, output dir {Dir}", openFileDialog.FileNames.Length, outputDirectory);
            }
        }

        private void btnSelectOutput_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                outputDirectory = folderBrowserDialog.SelectedPath;
                if (txtOutputDir != null)
                    txtOutputDir.Text = outputDirectory;
                Log($"输出目录：{outputDirectory}");
                AppLogger.Logger?.Information("Output directory set to {Dir}", outputDirectory);
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
                        bool skipFile = false;
                        // First check file accessibility before calling SplitFile
                        while (true)
                        {
                            try
                            {
                                using (var fs = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None)) { }
                                break; // accessible
                            }
                            catch (Exception ex)
                            {
                                // ask user to Retry or Skip
                                var dr = (DialogResult)Invoke(new Func<DialogResult>(() => MessageBox.Show(this,
                                    $"无法访问文件:\n{file}\n\n可能原因：文件正在被其他程序占用或没有权限。\n请关闭占用该文件的程序（例如 Excel），然后重试。\n\n详细信息：{ex.Message}",
                                    Text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning)));
                                if (dr == DialogResult.Retry)
                                    continue;
                                Invoke(new Action(() => Log($"跳过文件（无法访问）：{Path.GetFileName(file)}", error: true)));
                                skipFile = true;
                                break;
                            }
                        }

                        if (skipFile) continue;

                        try
                        {
                            SplitFile(file, maxRowsPerFile, headerRows, outputDirectory);
                            Invoke(new Action(() => Log($"拆分完成：{Path.GetFileName(file)}")));
                            AppLogger.Logger?.Information("Split completed for {File}", file);
                        }
                        catch (Exception ex)
                        {
                            // If it's a file lock/IO/permission issue, prompt user to Retry/Skip
                            if (IsFileLockException(ex))
                            {
                                var dr = (DialogResult)Invoke(new Func<DialogResult>(() => MessageBox.Show(this,
                                    $"处理文件时发生访问错误:\n{file}\n\n{ex.Message}\n\n请关闭占用该文件的程序（例如 Excel），然后重试。",
                                    Text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning)));
                                if (dr == DialogResult.Retry)
                                {
                                    // retry once by re-queueing this file
                                    // simple approach: try again immediately
                                    try
                                    {
                                        SplitFile(file, maxRowsPerFile, headerRows, outputDirectory);
                                        Invoke(new Action(() => Log($"拆分完成：{Path.GetFileName(file)}")));
                                    }
                                    catch (Exception rex)
                                    {
                                        Invoke(new Action(() => Log($"错误：{Path.GetFileName(file)} -> {rex.Message}")));
                                    }
                                }
                                else
                                {
                                    Invoke(new Action(() => Log($"跳过文件（被占用/无权限）：{Path.GetFileName(file)} -> {ex.Message}", error: true)));
                                }
                            }
                            else
                            {
                                Invoke(new Action(() => Log($"错误：{Path.GetFileName(file)} -> {ex.Message}")));
                            }
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

            // 打开 workbook（若被占用则提示重试或跳过）
            XLWorkbook wb;
            while (true)
            {
                try
                {
                    wb = new XLWorkbook(inputPath);
                    break;
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    var dr = (DialogResult)Invoke(new Func<DialogResult>(() => MessageBox.Show(this,
                        $"无法打开文件:\n{inputPath}\n\n可能原因：文件正在被其他程序占用或没有权限。\n请关闭占用该文件的程序（例如 Excel），然后重试。\n\n详细信息：{ex.Message}",
                        Text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning)));
                    if (dr == DialogResult.Retry)
                        continue;
                    Invoke(new Action(() => Log($"跳过文件（无法打开）：{Path.GetFileName(inputPath)}", error: true)));
                    return;
                }
                catch (Exception ex)
                {
                    Invoke(new Action(() => Log($"错误：{Path.GetFileName(inputPath)} -> {ex.Message}", error: true)));
                    return;
                }
            }

            using (wb)
            {
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
                if (remainder == 0)
                {
                    var destName = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(inputPath)}_sheets{Path.GetExtension(inputPath)}");
                    if (TrySaveWorkbook(wb, destName))
                        Invoke(new Action(() => Log($"无需拆分或数据量不足，已复制为：{Path.GetFileName(destName)}")));
                    else
                        Invoke(new Action(() => Log($"保存失败：{Path.GetFileName(destName)}（可能被占用或无权限）", error: true)));
                    return;
                }

                // 是否拆分成多个文件（复选框控制）
                bool splitToFiles = chkSplitToFiles?.Checked ?? false;

                // 解析分组列（逗号分隔），如果为空或无效则使用原有按行拆分逻辑
                var groupColsText = txtGroupCols?.Text?.Trim();
                int[] groupCols = Array.Empty<int>();
                if (!string.IsNullOrEmpty(groupColsText))
                {
                    try
                    {
                        groupCols = groupColsText.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => ToColumnNumber(s.Trim().ToUpperInvariant()))
                            .Where(v => v > 0 && v <= lastColUsed)
                            .Distinct()
                            .OrderBy(v => v)
                            .ToArray();
                    }
                    catch
                    {
                        groupCols = Array.Empty<int>();
                    }
                }

                double[] columnWidths = new double[lastColUsed + 1];
                for (int c = 1; c <= lastColUsed; c++)
                    columnWidths[c] = ws.Column(c).Width;

                int partIndex = 1;

                if (groupCols.Length == 0)
                {
                    // 原有按行拆分逻辑
                    for (int offset = 0; offset < remainder; offset += maxRowsPerFile)
                    {
                        int chunkStart = dataStart + offset;
                        int chunkEnd = Math.Min(chunkStart + maxRowsPerFile - 1, lastRowUsed);

                        if (splitToFiles)
                        {
                            // 为每个 chunk 新建 workbook 并保存为单独文件
                            using var newWb = new XLWorkbook();
                            var newWs = newWb.Worksheets.Add(ws.Name);
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

                            string destFile = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(inputPath)}_part{partIndex}{Path.GetExtension(inputPath)}");
                            if (TrySaveWorkbook(newWb, destFile))
                                Invoke(new Action(() => Log($"已生成文件：{Path.GetFileName(destFile)} (行 {chunkStart}-{chunkEnd}，含头部 {headerRows} 行)")));
                            else
                                Invoke(new Action(() => Log($"保存失败：{Path.GetFileName(destFile)}（可能被占用或无权限）", error: true)));

                            partIndex++;
                        }
                        else
                        {
                            // 在同一 workbook 中新增 sheet
                            string baseName = ws.Name + "_part" + partIndex;
                            string newSheetName = baseName;
                            int tryIndex = 1;
                            while (wb.Worksheets.Any(s => s.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)))
                            {
                                tryIndex++;
                                newSheetName = baseName + "(" + tryIndex + ")";
                            }

                            var newWs = wb.Worksheets.Add(newSheetName);
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
                    }
                }
                else
                {
                    // 按分组列分组，保持组内行不拆分，同组数据必须在同一个 sheet
                    var groups = new Dictionary<string, List<int>>();
                    var groupOrder = new List<string>();
                    for (int r = dataStart; r <= lastRowUsed; r++)
                    {
                        var keyParts = new string[groupCols.Length];
                        for (int i = 0; i < groupCols.Length; i++)
                        {
                            var col = groupCols[i];
                            keyParts[i] = ws.Cell(r, col).GetString();
                        }
                        var key = string.Join("\u001F", keyParts);
                        if (!groups.ContainsKey(key))
                        {
                            groups[key] = new List<int>();
                            groupOrder.Add(key);
                        }
                        groups[key].Add(r);
                    }

                    // 按组打包到 chunks，每个 chunk 包含若干完整的组，且数据行数不超过 maxRowsPerFile（若单组超限则允许超限但不拆分）
                    var chunks = new List<List<int>>();
                    var current = new List<int>();
                    int currentCount = 0;
                    foreach (var key in groupOrder)
                    {
                        var rows = groups[key];
                        int size = rows.Count;
                        if (currentCount + size > maxRowsPerFile)
                        {
                            if (currentCount == 0)
                            {
                                // 单个组超出 max，仍放在单独 sheet 中
                                chunks.Add(new List<int>(rows));
                            }
                            else
                            {
                                // 先保存当前 chunk，再开始新 chunk
                                chunks.Add(new List<int>(current));
                                current.Clear();
                                currentCount = 0;
                                // 将当前组放入新 chunk
                                if (size > maxRowsPerFile)
                                {
                                    chunks.Add(new List<int>(rows));
                                }
                                else
                                {
                                    current.AddRange(rows);
                                    currentCount += size;
                                }
                            }
                        }
                        else
                        {
                            current.AddRange(rows);
                            currentCount += size;
                        }
                    }
                    if (current.Count > 0)
                        chunks.Add(new List<int>(current));

                    // 为每个 chunk 生成 sheet 或文件
                    foreach (var chunk in chunks)
                    {
                        if (splitToFiles)
                        {
                            // save each chunk as a separate workbook
                            using var newWb = new XLWorkbook();
                            var newWs = newWb.Worksheets.Add(ws.Name);
                            for (int c = 1; c <= lastColUsed; c++)
                                newWs.Column(c).Width = columnWidths[c];

                            int destRow = 1;
                            // copy headers
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

                            int chunkStart = int.MaxValue;
                            int chunkEnd = int.MinValue;
                            foreach (var r in chunk)
                            {
                                chunkStart = Math.Min(chunkStart, r);
                                chunkEnd = Math.Max(chunkEnd, r);
                                for (int c = 1; c <= lastColUsed; c++)
                                {
                                    var srcCell = ws.Cell(r, c);
                                    var dstCell = newWs.Cell(destRow, c);
                                    dstCell.Value = srcCell.Value;
                                    dstCell.Style = srcCell.Style;
                                }
                                destRow++;
                            }

                            string destFile = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(inputPath)}_part{partIndex}{Path.GetExtension(inputPath)}");
                            if (TrySaveWorkbook(newWb, destFile))
                                Invoke(new Action(() => Log($"已生成文件：{Path.GetFileName(destFile)} (行 {chunkStart}-{chunkEnd}，含头部 {headerRows} 行，组数 {chunk.Count})")));
                            else
                                Invoke(new Action(() => Log($"保存失败：{Path.GetFileName(destFile)}（可能被占用或无权限）", error: true)));

                            partIndex++;
                        }
                        else
                        {
                            string baseName = ws.Name + "_part" + partIndex;
                            string newSheetName = baseName;
                            int tryIndex = 1;
                            while (wb.Worksheets.Any(s => s.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)))
                            {
                                tryIndex++;
                                newSheetName = baseName + "(" + tryIndex + ")";
                            }

                            var newWs = wb.Worksheets.Add(newSheetName);
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

                            int chunkStart = int.MaxValue;
                            int chunkEnd = int.MinValue;
                            foreach (var r in chunk)
                            {
                                chunkStart = Math.Min(chunkStart, r);
                                chunkEnd = Math.Max(chunkEnd, r);
                                for (int c = 1; c <= lastColUsed; c++)
                                {
                                    var srcCell = ws.Cell(r, c);
                                    var dstCell = newWs.Cell(destRow, c);
                                    dstCell.Value = srcCell.Value;
                                    dstCell.Style = srcCell.Style;
                                }
                                destRow++;
                            }

                            Invoke(new Action(() => Log($"已生成 sheet：{newSheetName} (行 {chunkStart}-{chunkEnd}，含头部 {headerRows} 行，组数 {chunk.Count})")));
                            partIndex++;
                        }
                    }
                }

                // 保存为一个新文件，保留原文件不变
                string outFile = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(inputPath)}_sheets{Path.GetExtension(inputPath)}");
                if (TrySaveWorkbook(wb, outFile))
                    Invoke(new Action(() => Log($"已保存带分片的文件：{Path.GetFileName(outFile)}")));
                else
                    Invoke(new Action(() => Log($"保存失败：{Path.GetFileName(outFile)}（可能被占用或无权限）")));
            }
        }


        private bool TrySaveWorkbook(XLWorkbook wb, string path)
        {
            try
            {
                wb.SaveAs(path);
                return true;
            }
            catch (UnauthorizedAccessException uex)
            {
                Invoke(new Action(() => MessageBox.Show(this, $"无法保存文件:\n{path}\n\n原因：没有写入权限或文件正在被其他程序占用。请关闭可能正在使用该文件的程序，或选择其他输出目录后重试。\n\n详细信息：{uex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                return false;
            }
            catch (IOException ioex)
            {
                Invoke(new Action(() => MessageBox.Show(this, $"无法保存文件:\n{path}\n\n原因：文件可能正在被其他程序占用。请关闭对应的程序（例如 Excel），或选择其他输出目录后重试。\n\n详细信息：{ioex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                return false;
            }
            catch (Exception ex)
            {
                Invoke(new Action(() => MessageBox.Show(this, $"保存文件时发生错误：{ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)));
                return false;
            }
        }

        private void Log(string message, bool warning = false, bool error = false)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Log(message)));
                return;
            }
            var line = $"{DateTime.Now:HH:mm:ss} {message}";
            // append to RichTextBox
            lstLog.AppendText(line + Environment.NewLine);
            lstLog.SelectionStart = lstLog.Text.Length;
            lstLog.ScrollToCaret();
            try
            {
                if (error)
                {
                    AppLogger.Logger?.Error(line);
                }
                else if (warning)
                {
                    AppLogger.Logger?.Warning(line);
                }
                else
                {
                    AppLogger.Logger?.Information(line);

                }
            }
            catch
            {
                // ignore logging errors
            }
        }

        private void menuCopy_Click(object? sender, EventArgs e)
        {
            var lb = ResolveListBoxFromMenuSender(sender);
            if (lb == null) return;
            if (lb.SelectedItems.Count == 0) return;
            CopyItemsToClipboard(lb.SelectedItems.Cast<object>().Select(o => o?.ToString() ?? string.Empty));
        }

        private void menuCopyAll_Click(object? sender, EventArgs e)
        {
            var lb = ResolveListBoxFromMenuSender(sender);
            if (lb == null) return;
            CopyItemsToClipboard(lb.Items.Cast<object>().Select(o => o?.ToString() ?? string.Empty));
        }

        private void lstList_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                var lb = sender as ListBox;
                if (lb == null) return;
                if (lb.SelectedItems.Count > 0)
                {
                    CopyItemsToClipboard(lb.SelectedItems.Cast<object>().Select(o => o?.ToString() ?? string.Empty));
                }
                else
                {
                    CopyItemsToClipboard(lb.Items.Cast<object>().Select(o => o?.ToString() ?? string.Empty));
                }
                e.Handled = true;
            }
        }

        private ListBox? ResolveListBoxFromMenuSender(object? sender)
        {
            // Try to get ContextMenuStrip from sender
            if (sender is ToolStripMenuItem tsi)
            {
                if (tsi.Owner is ContextMenuStrip cms)
                {
                    if (cms.SourceControl is ListBox lb)
                        return lb;
                }
            }
            // Fallback: if one of our listboxes has focus, use it
            if (lstFiles.Focused) return lstFiles;
            // lstLog is now RichTextBox; if focused, prefer lstFiles fallback
            if (lstLog.Focused) return lstFiles;
            // Last resort: return lstFiles if it has items
            if (lstFiles.Items.Count > 0) return lstFiles;
            // RichTextBox doesn't have Items; skip
            return null;
        }

        private void CopyItemsToClipboard(IEnumerable<string> lines)
        {
            var text = string.Join(Environment.NewLine, lines);
            try
            {
                Clipboard.SetText(text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"复制到剪贴板失败：{ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static bool IsFileLockException(Exception ex)
        {
            if (ex is IOException) return true;
            if (ex is UnauthorizedAccessException) return true;
            var msg = ex.Message ?? string.Empty;
            if (msg.IndexOf("being used by another process", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void chkSplitToFiles_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
