using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using DwgOfficeToolbox.App.Models;
using DwgOfficeToolbox.App.Services;

namespace DwgOfficeToolbox.App;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<InputItem> _inputs = new();
    private readonly ObservableCollection<MatchResult> _results = new();
    private readonly ObservableCollection<string> _logs = new();
    private readonly ICollectionView _resultsView;
    private readonly AppSettings _settings;
    private CancellationTokenSource? _cts;
    private bool _isRunning;
    private bool _isUpdatingFilters;
    private List<MatchResult> _allResults = new();
    private readonly object _logFileLock = new();
    private string? _logFilePath;

    public MainWindow()
    {
        InitializeComponent();
        _settings = SettingsManager.Load();

        InputList.ItemsSource = _inputs;
        LogList.ItemsSource = _logs;

        _resultsView = CollectionViewSource.GetDefaultView(_results);
        _resultsView.Filter = FilterResults;
        ResultsGrid.ItemsSource = _resultsView;

        InitializeFilters();
        UpdateStats();
    }

    private void InitializeFilters()
    {
        _isUpdatingFilters = true;
        FileFilterCombo.ItemsSource = new List<string> { "全部" };
        TypeFilterCombo.ItemsSource = new List<string> { "全部" };
        LayerFilterCombo.ItemsSource = new List<string> { "全部" };
        KeywordFilterCombo.ItemsSource = new List<string> { "全部" };
        FileFilterCombo.SelectedIndex = 0;
        TypeFilterCombo.SelectedIndex = 0;
        LayerFilterCombo.SelectedIndex = 0;
        KeywordFilterCombo.SelectedIndex = 0;
        _isUpdatingFilters = false;
    }

    private void AddFiles_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "DWG 文件 (*.dwg)|*.dwg",
            Multiselect = true
        };
        if (dialog.ShowDialog() == true)
        {
            foreach (var file in dialog.FileNames)
            {
                AddInput(file, isFolder: false);
            }
        }
    }

    private void AddFolder_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        var result = dialog.ShowDialog();
        if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
        {
            AddInput(dialog.SelectedPath, isFolder: true);
        }
    }

    private void RemoveSelected_Click(object sender, RoutedEventArgs e)
    {
        var selected = InputList.SelectedItems.Cast<InputItem>().ToList();
        foreach (var item in selected)
        {
            _inputs.Remove(item);
        }
    }

    private void ClearInput_Click(object sender, RoutedEventArgs e)
    {
        _inputs.Clear();
    }

    private void AddInput(string path, bool isFolder)
    {
        var full = Path.GetFullPath(path);
        if (_inputs.Any(i => string.Equals(i.Path, full, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }
        _inputs.Add(new InputItem
        {
            Path = full,
            IsFolder = isFolder,
            Type = isFolder ? "文件夹" : "文件"
        });
    }

    private async void Start_Click(object sender, RoutedEventArgs e)
    {
        if (_isRunning)
        {
            return;
        }

        if (_inputs.Count == 0)
        {
            System.Windows.MessageBox.Show("请先选择 DWG 文件或文件夹。");
            return;
        }

        var odaPath = ResolveOdaPath();
        if (!File.Exists(odaPath))
        {
            System.Windows.MessageBox.Show($"未找到 ODAFileConverter.exe，请检查设置路径：{odaPath}");
            return;
        }

        _cts = new CancellationTokenSource();
        _isRunning = true;
        ToggleUi(true);
        ClearResults();
        _logFilePath = CreateLogFile();
        Log($"日志文件：{_logFilePath}");
        Log("开始处理...");

        try
        {
            var keywords = ParseKeywords(KeywordsTextBox.Text);
            var inputFolders = _inputs.Where(i => i.IsFolder).Select(i => i.Path).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var inputFiles = _inputs.Where(i => !i.IsFolder).Select(i => i.Path).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            inputFiles = inputFiles.Where(f => !IsCoveredByFolder(f, inputFolders)).ToList();

            ProgressText.Text = "准备文件清单...";
            var scanTargets = await Task.Run(
                () => BuildScanTargets(inputFolders, inputFiles, _settings.OutputFolderName),
                _cts.Token);

            if (_settings.ClearOutputBeforeConvert)
            {
                ClearOutputRoots(scanTargets);
            }

            var compareResult = await CompareTargetsAsync(scanTargets, _cts.Token);
            Log($"待转换：{compareResult.Pending.Count}，已存在跳过：{compareResult.Skipped}");

            await RunConversionsAsync(compareResult.Pending, odaPath, _cts.Token);
            await RunScanAsync(scanTargets, keywords, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            Log("已取消。");
        }
        catch (Exception ex)
        {
            Log($"发生错误：{ex.Message}");
        }
        finally
        {
            _isRunning = false;
            ToggleUi(false);
            _cts = null;
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
    }

    private void ExportCsv_Click(object sender, RoutedEventArgs e)
    {
        if (_allResults.Count == 0)
        {
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV 文件 (*.csv)|*.csv",
            FileName = $"扫描结果_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dialog.ShowDialog() == true)
        {
            ExportService.ExportCsv(dialog.FileName, _allResults);
            Log($"CSV 已导出：{dialog.FileName}");
        }
    }

    private void ExportXlsx_Click(object sender, RoutedEventArgs e)
    {
        if (_allResults.Count == 0)
        {
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel 文件 (*.xlsx)|*.xlsx",
            FileName = $"扫描结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
        };
        if (dialog.ShowDialog() == true)
        {
            ExportService.ExportXlsx(dialog.FileName, _allResults);
            Log($"Excel 已导出：{dialog.FileName}");
        }
    }

    private void OpenDwgLink_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Documents.Hyperlink link)
        {
            return;
        }

        if (link.DataContext is not MatchResult result)
        {
            return;
        }

        var path = result.DwgPath;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            System.Windows.MessageBox.Show("未找到对应 DWG 文件。");
            return;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            };
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"打开文件失败：{ex.Message}");
        }
    }

    private void FilterChanged(object sender, SelectionChangedEventArgs e)
    {
        RefreshFilters();
    }

    private void ContentFilterTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        RefreshFilters();
    }

    private void RefreshFilters()
    {
        if (_isUpdatingFilters)
        {
            return;
        }
        _resultsView.Refresh();
        UpdateStats();
    }

    private bool FilterResults(object obj)
    {
        if (obj is not MatchResult result)
        {
            return false;
        }

        var file = FileFilterCombo.SelectedItem as string ?? "全部";
        var type = TypeFilterCombo.SelectedItem as string ?? "全部";
        var layer = LayerFilterCombo.SelectedItem as string ?? "全部";
        var keyword = KeywordFilterCombo.SelectedItem as string ?? "全部";
        var content = ContentFilterTextBox.Text?.Trim() ?? string.Empty;

        if (file != "全部" && !string.Equals(result.FileName, file, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (type != "全部" && !string.Equals(result.ObjectType, type, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (layer != "全部" && !string.Equals(result.Layer, layer, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (keyword != "全部" && !string.Equals(result.Keyword, keyword, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(content) &&
            !result.Content.Contains(content, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private List<string> ParseKeywords(string raw)
    {
        return raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(k => !string.IsNullOrWhiteSpace(k))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsCoveredByFolder(string filePath, List<string> folders)
    {
        foreach (var folder in folders)
        {
            if (IsSubPath(folder, filePath))
            {
                return true;
            }
        }
        return false;
    }

    private static bool IsSubPath(string parent, string child)
    {
        var parentFull = Path.GetFullPath(parent)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var childFull = Path.GetFullPath(child);
        return childFull.StartsWith(parentFull, StringComparison.OrdinalIgnoreCase);
    }

    private List<DwgScanTarget> BuildScanTargets(List<string> folders, List<string> files, string outputFolderName)
    {
        var targets = new List<DwgScanTarget>();

        foreach (var folder in folders)
        {
            var outputRoot = Path.Combine(folder, outputFolderName);
            foreach (var dwg in Directory.EnumerateFiles(folder, _settings.InputFilter, SearchOption.AllDirectories))
            {
                var rel = Path.GetRelativePath(folder, dwg);
                var dxfRel = Path.ChangeExtension(rel, ".dxf");
                var dxfPath = Path.Combine(outputRoot, dxfRel);
                targets.Add(new DwgScanTarget { DwgPath = dwg, DxfPath = dxfPath, OutputRoot = outputRoot });
            }
        }

        foreach (var file in files)
        {
            var dir = Path.GetDirectoryName(file) ?? ".";
            var outputRoot = Path.Combine(dir, outputFolderName);
            var dxfPath = Path.Combine(outputRoot, Path.ChangeExtension(Path.GetFileName(file), ".dxf"));
            targets.Add(new DwgScanTarget { DwgPath = file, DxfPath = dxfPath, OutputRoot = outputRoot });
        }

        return targets;
    }

    private async Task RunConversionsAsync(List<DwgScanTarget> targets, string odaPath, CancellationToken ct)
    {
        if (targets.Count == 0)
        {
            Log("没有需要转换的文件。");
            ProgressText.Text = "无需转换";
            return;
        }

        ProgressBar.Value = 0;
        ProgressBar.Maximum = targets.Count;
        var processed = 0;
        var maxConcurrency = Math.Max(1, _settings.MaxConvertConcurrency);
        using var semaphore = new SemaphoreSlim(maxConcurrency);

        var tasks = targets.Select(async target =>
        {
            await semaphore.WaitAsync(ct);
            try
            {
                ct.ThrowIfCancellationRequested();
                var inputFolder = Path.GetDirectoryName(target.DwgPath) ?? ".";
                var outputFolder = Path.GetDirectoryName(target.DxfPath) ?? target.OutputRoot;
                var inputFilter = Path.GetFileName(target.DwgPath);
                Directory.CreateDirectory(outputFolder);

                Log($"转换：{target.DwgPath}");
                var exitCode = await OdaRunner.ConvertAsync(
                    odaPath,
                    inputFolder,
                    outputFolder,
                    _settings.OutputVersion,
                    _settings.OutputFormat,
                    inputFilter,
                    recurse: false,
                    audit: true,
                    ct,
                    Log);

                if (exitCode != 0)
                {
                    Log($"ODA 转换返回码：{exitCode}");
                }
            }
            finally
            {
                var current = Interlocked.Increment(ref processed);
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Value = current;
                    ProgressText.Text = $"转换中 ({current}/{targets.Count})";
                });
                semaphore.Release();
            }
        }).ToList();

        await Task.WhenAll(tasks);
    }

    private async Task<CompareResult> CompareTargetsAsync(List<DwgScanTarget> targets, CancellationToken ct)
    {
        CompareProgressBar.Value = 0;
        CompareProgressBar.Maximum = targets.Count == 0 ? 1 : targets.Count;
        CompareProgressText.Text = targets.Count == 0 ? "0/0" : $"0/{targets.Count}";

        return await Task.Run(() =>
        {
            var pending = new List<DwgScanTarget>();
            var total = targets.Count;
            var lastUpdate = 0;
            var step = Math.Max(1, total / 200);

            for (var i = 0; i < total; i++)
            {
                ct.ThrowIfCancellationRequested();
                var target = targets[i];
                if (NeedsConvert(target))
                {
                    pending.Add(target);
                }

                var current = i + 1;
                if (current - lastUpdate >= step || current == total)
                {
                    lastUpdate = current;
                    Dispatcher.Invoke(() =>
                    {
                        CompareProgressBar.Value = current;
                        CompareProgressText.Text = $"{current}/{total}";
                    });
                }
            }

            Dispatcher.Invoke(() =>
            {
                CompareProgressText.Text = $"比对完成 {total}/{total}";
            });

            return new CompareResult(pending, total - pending.Count);
        }, ct);
    }

    private async Task RunScanAsync(List<DwgScanTarget> targets, List<string> keywords, CancellationToken ct)
    {
        if (_settings.EnableTextCache)
        {
            await RunScanWithCacheAsync(targets, keywords, ct);
            return;
        }

        ProgressBar.Value = 0;
        ProgressBar.Maximum = targets.Count == 0 ? 1 : targets.Count;
        ProgressText.Text = "开始扫描...";

        var bag = new ConcurrentBag<MatchResult>();
        var processed = 0;
        var maxConcurrency = Math.Max(1, Math.Min(Environment.ProcessorCount, _settings.MaxParseConcurrency));
        var semaphore = new SemaphoreSlim(maxConcurrency);
        var maxKeywordLen = keywords.Count == 0 ? 0 : keywords.Max(k => k.Length);

        var tasks = targets.Select(async target =>
        {
            await semaphore.WaitAsync(ct);
            try
            {
                ct.ThrowIfCancellationRequested();
                if (!File.Exists(target.DxfPath))
                {
                    Log($"未找到 DXF：{target.DxfPath}");
                    return;
                }

                var fileInfo = new FileInfo(target.DxfPath);
                var thresholdBytes = _settings.LargeFileThresholdMB * 1024L * 1024L;
                if (fileInfo.Length > thresholdBytes)
                {
                    var plainResults = KeywordScanner.ScanPlainText(
                        target.DxfPath,
                        target.DwgPath,
                        keywords,
                        ks => MatchKeywordsStream(target.DxfPath, ks, maxKeywordLen));
                    foreach (var r in plainResults)
                    {
                        bag.Add(r);
                    }
                    return;
                }

                try
                {
                    var items = DxfTextExtractor.Extract(target.DxfPath);
                    if (items.Count == 0 && keywords.Count > 0)
                    {
                        Log($"解析未提取到文字，改用纯文本：{Path.GetFileName(target.DxfPath)}");
                        var plainResults = KeywordScanner.ScanPlainText(
                            target.DxfPath,
                            target.DwgPath,
                            keywords,
                            ks => MatchKeywordsStream(target.DxfPath, ks, maxKeywordLen));
                        foreach (var r in plainResults)
                        {
                            bag.Add(r);
                        }
                    }
                    else
                    {
                        var structured = KeywordScanner.ScanStructured(target.DxfPath, target.DwgPath, items, keywords);
                        foreach (var r in structured)
                        {
                            bag.Add(r);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"解析失败，改用纯文本：{Path.GetFileName(target.DxfPath)}，{ex.Message}");
                    var plainResults = KeywordScanner.ScanPlainText(
                        target.DxfPath,
                        target.DwgPath,
                        keywords,
                        ks => MatchKeywordsStream(target.DxfPath, ks, maxKeywordLen));
                    foreach (var r in plainResults)
                    {
                        bag.Add(r);
                    }
                }
            }
            finally
            {
                var current = Interlocked.Increment(ref processed);
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Value = current;
                    ProgressText.Text = $"扫描中 ({current}/{targets.Count})";
                });
                semaphore.Release();
            }
        }).ToList();

        await Task.WhenAll(tasks);

        _allResults = bag.OrderBy(r => r.FileName, StringComparer.OrdinalIgnoreCase).ToList();
        _results.Clear();
        foreach (var r in _allResults)
        {
            _results.Add(r);
        }

        UpdateFilterOptions();
        RefreshFilters();

        ExportCsvButton.IsEnabled = _allResults.Count > 0;
        ExportXlsxButton.IsEnabled = _allResults.Count > 0;

        Log($"扫描完成，结果数：{_allResults.Count}");
        ProgressText.Text = "扫描完成";

        if (!_settings.KeepConvertedDxf)
        {
            foreach (var outputRoot in targets.Select(t => t.OutputRoot).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    if (Directory.Exists(outputRoot))
                    {
                        Directory.Delete(outputRoot, recursive: true);
                        Log($"已清理输出目录：{outputRoot}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"清理输出目录失败：{outputRoot}，{ex.Message}");
                }
            }
        }
    }

    private HashSet<string> MatchKeywordsStream(string filePath, IReadOnlyList<string> keywords, int maxKeywordLen)
    {
        var matched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (keywords.Count == 0)
        {
            return matched;
        }

        var encoding = DetectEncoding(filePath);
        using var reader = new StreamReader(filePath, encoding, detectEncodingFromByteOrderMarks: true);
        var buffer = new char[8192];
        var carry = string.Empty;
        int read;
        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
        {
            var chunk = carry + new string(buffer, 0, read);
            foreach (var keyword in keywords)
            {
                if (!matched.Contains(keyword) &&
                    chunk.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    matched.Add(keyword);
                }
            }

            if (matched.Count == keywords.Count)
            {
                break;
            }

            if (maxKeywordLen > 1 && chunk.Length >= maxKeywordLen - 1)
            {
                carry = chunk.Substring(chunk.Length - (maxKeywordLen - 1));
            }
            else
            {
                carry = chunk;
            }
        }

        return matched;
    }

    private async Task RunScanWithCacheAsync(List<DwgScanTarget> targets, List<string> keywords, CancellationToken ct)
    {
        ProgressBar.Value = 0;
        ProgressBar.Maximum = targets.Count == 0 ? 1 : targets.Count;
        ProgressText.Text = "更新缓存...";

        var bag = new ConcurrentBag<MatchResult>();
        var grouped = targets.GroupBy(t => t.OutputRoot, StringComparer.OrdinalIgnoreCase).ToList();
        var totalTargets = targets.Count;
        var processedBase = 0;
        var cachedTotal = 0;
        var skippedTotal = 0;
        var failedTotal = 0;
        var allPlainTargets = new List<DwgScanTarget>();

        foreach (var group in grouped)
        {
            ct.ThrowIfCancellationRequested();
            var outputRoot = group.Key;
            Directory.CreateDirectory(outputRoot);
            var dbPath = Path.Combine(outputRoot, ".cadtext_cache.db");
            var cacheService = new TextCacheService(dbPath);

            var groupTargets = group.ToList();
            var groupResult = await Task.Run(() =>
                cacheService.UpdateCache(
                    groupTargets,
                    _settings.LargeFileThresholdMB,
                    Log,
                    (current, total) =>
                    {
                        var absolute = processedBase + current;
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBar.Value = absolute;
                            ProgressText.Text = $"缓存更新 ({absolute}/{totalTargets})";
                        });
                    },
                    ct), ct);

            cachedTotal += groupResult.CachedCount;
            skippedTotal += groupResult.SkippedCount;
            failedTotal += groupResult.FailedCount;
            allPlainTargets.AddRange(groupResult.PlainTextTargets);
            processedBase += groupTargets.Count;
            SetHiddenIfExists(dbPath);
        }

        Log($"缓存更新完成：新增 {cachedTotal}，跳过 {skippedTotal}，未缓存 {failedTotal}");

        ProgressBar.Value = 0;
        ProgressBar.Maximum = 1;
        ProgressText.Text = "读取缓存结果...";

        foreach (var group in grouped)
        {
            var outputRoot = group.Key;
            var dbPath = Path.Combine(outputRoot, ".cadtext_cache.db");
            var cacheService = new TextCacheService(dbPath);
            var cachedResults = await Task.Run(() => cacheService.Query(group.ToList(), keywords), ct);
            foreach (var r in cachedResults)
            {
                bag.Add(r);
            }
        }

        if (allPlainTargets.Count > 0)
        {
            await RunPlainTextScanAsync(allPlainTargets, keywords, bag, ct);
        }

        _allResults = bag.OrderBy(r => r.FileName, StringComparer.OrdinalIgnoreCase).ToList();
        _results.Clear();
        foreach (var r in _allResults)
        {
            _results.Add(r);
        }

        UpdateFilterOptions();
        RefreshFilters();

        ExportCsvButton.IsEnabled = _allResults.Count > 0;
        ExportXlsxButton.IsEnabled = _allResults.Count > 0;

        Log($"扫描完成，结果数：{_allResults.Count}");
        ProgressText.Text = "扫描完成";

        if (!_settings.KeepConvertedDxf)
        {
            foreach (var outputRoot in targets.Select(t => t.OutputRoot).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    if (Directory.Exists(outputRoot))
                    {
                        Directory.Delete(outputRoot, recursive: true);
                        Log($"已清理输出目录：{outputRoot}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"清理输出目录失败：{outputRoot}，{ex.Message}");
                }
            }
        }
    }

    private static void SetHiddenIfExists(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return;
            }

            var attrs = File.GetAttributes(path);
            if ((attrs & FileAttributes.Hidden) == 0)
            {
                File.SetAttributes(path, attrs | FileAttributes.Hidden);
            }
        }
        catch
        {
            // ignore
        }
    }

    private async Task RunPlainTextScanAsync(
        List<DwgScanTarget> targets,
        List<string> keywords,
        ConcurrentBag<MatchResult> bag,
        CancellationToken ct)
    {
        if (targets.Count == 0)
        {
            return;
        }

        ProgressBar.Value = 0;
        ProgressBar.Maximum = targets.Count;
        ProgressText.Text = "纯文本扫描中...";

        var processed = 0;
        var maxKeywordLen = keywords.Count == 0 ? 0 : keywords.Max(k => k.Length);
        var maxConcurrency = Math.Max(1, Math.Min(Environment.ProcessorCount, _settings.MaxParseConcurrency));
        using var semaphore = new SemaphoreSlim(maxConcurrency);

        var tasks = targets.Select(async target =>
        {
            await semaphore.WaitAsync(ct);
            try
            {
                ct.ThrowIfCancellationRequested();
                if (!File.Exists(target.DxfPath))
                {
                    return;
                }

                var plainResults = KeywordScanner.ScanPlainText(
                    target.DxfPath,
                    target.DwgPath,
                    keywords,
                    ks => MatchKeywordsStream(target.DxfPath, ks, maxKeywordLen));
                foreach (var r in plainResults)
                {
                    bag.Add(r);
                }
            }
            finally
            {
                var current = Interlocked.Increment(ref processed);
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Value = current;
                    ProgressText.Text = $"纯文本扫描 ({current}/{targets.Count})";
                });
                semaphore.Release();
            }
        }).ToList();

        await Task.WhenAll(tasks);
    }

    private static Encoding DetectEncoding(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        Span<byte> bom = stackalloc byte[4];
        var read = stream.Read(bom);
        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
        {
            return Encoding.UTF8;
        }

        if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
        {
            return Encoding.Unicode;
        }

        if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
        {
            return Encoding.BigEndianUnicode;
        }

        var sampleLength = (int)Math.Min(4096, stream.Length);
        if (sampleLength <= 0)
        {
            return Encoding.UTF8;
        }

        stream.Position = 0;
        var sample = new byte[sampleLength];
        var count = stream.Read(sample, 0, sample.Length);
        var zeroEven = 0;
        var zeroOdd = 0;
        for (var i = 0; i < count; i++)
        {
            if (sample[i] == 0)
            {
                if (i % 2 == 0)
                {
                    zeroEven++;
                }
                else
                {
                    zeroOdd++;
                }
            }
        }

        var zeroTotal = zeroEven + zeroOdd;
        if (zeroTotal > count / 10)
        {
            return zeroEven >= zeroOdd ? Encoding.Unicode : Encoding.BigEndianUnicode;
        }

        return Encoding.UTF8;
    }

    private void UpdateFilterOptions()
    {
        _isUpdatingFilters = true;
        FileFilterCombo.ItemsSource = BuildOptions(_results.Select(r => r.FileName));
        TypeFilterCombo.ItemsSource = BuildOptions(_results.Select(r => r.ObjectType));
        LayerFilterCombo.ItemsSource = BuildOptions(_results.Select(r => r.Layer));
        KeywordFilterCombo.ItemsSource = BuildOptions(_results.Select(r => r.Keyword));
        FileFilterCombo.SelectedIndex = 0;
        TypeFilterCombo.SelectedIndex = 0;
        LayerFilterCombo.SelectedIndex = 0;
        KeywordFilterCombo.SelectedIndex = 0;
        _isUpdatingFilters = false;
    }

    private static List<string> BuildOptions(IEnumerable<string> values)
    {
        var list = values.Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
            .ToList();
        list.Insert(0, "全部");
        return list;
    }

    private void ToggleUi(bool isRunning)
    {
        StartButton.IsEnabled = !isRunning;
        CancelButton.IsEnabled = isRunning;
        AddFilesButton.IsEnabled = !isRunning;
        AddFolderButton.IsEnabled = !isRunning;
        RemoveSelectedButton.IsEnabled = !isRunning;
        ClearInputButton.IsEnabled = !isRunning;
        ExportCsvButton.IsEnabled = !isRunning && _allResults.Count > 0;
        ExportXlsxButton.IsEnabled = !isRunning && _allResults.Count > 0;
    }

    private void ClearResults()
    {
        _allResults = new List<MatchResult>();
        _results.Clear();
        InitializeFilters();
        UpdateStats();
        ExportCsvButton.IsEnabled = false;
        ExportXlsxButton.IsEnabled = false;
        ProgressBar.Value = 0;
        ProgressText.Text = string.Empty;
        CompareProgressBar.Value = 0;
        CompareProgressText.Text = string.Empty;
    }

    private void UpdateStats()
    {
        var viewItems = _resultsView.Cast<MatchResult>().ToList();
        var fileCount = viewItems.Select(r => r.FileName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        var keywordCount = viewItems.Select(r => r.Keyword)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        StatsText.Text = $"结果数：{viewItems.Count}  文件数：{fileCount}  关键词数：{keywordCount}";
    }

    private void Log(string message)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {message}";
        Dispatcher.Invoke(() =>
        {
            _logs.Add(line);
            if (_logs.Count > 2000)
            {
                _logs.RemoveAt(0);
            }
            LogList.ScrollIntoView(_logs.Last());
        });

        if (!string.IsNullOrWhiteSpace(_logFilePath))
        {
            lock (_logFileLock)
            {
                try
                {
                    File.AppendAllText(_logFilePath, line + Environment.NewLine);
                }
                catch
                {
                    // ignore log failures
                }
            }
        }
    }

    private static bool NeedsConvert(DwgScanTarget target)
    {
        if (!File.Exists(target.DxfPath))
        {
            return true;
        }

        var dxfTime = File.GetLastWriteTimeUtc(target.DxfPath);
        var dwgTime = File.GetLastWriteTimeUtc(target.DwgPath);
        return dxfTime < dwgTime;
    }

    private void ClearOutputRoots(List<DwgScanTarget> targets)
    {
        foreach (var outputRoot in targets.Select(t => t.OutputRoot).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                if (Directory.Exists(outputRoot))
                {
                    Directory.Delete(outputRoot, recursive: true);
                    Log($"已清理输出目录：{outputRoot}");
                }
            }
            catch (Exception ex)
            {
                Log($"清理输出目录失败：{outputRoot}，{ex.Message}");
            }
        }
    }

    private string CreateLogFile()
    {
        var dir = Path.Combine(SettingsManager.AppDataDir, "logs");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, $"run_{DateTime.Now:yyyyMMdd_HHmmss}.log");
        File.AppendAllText(path, $"Start {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}");
        return path;
    }

    private string ResolveOdaPath()
    {
        if (File.Exists(_settings.OdaExePath))
        {
            return _settings.OdaExePath;
        }

        var defaultPath = Path.Combine(AppContext.BaseDirectory, "ODAFileConverter", "ODAFileConverter.exe");
        if (File.Exists(defaultPath))
        {
            _settings.OdaExePath = defaultPath;
            SettingsManager.Save(_settings);
            Log($"已自动修复 ODA 路径为：{defaultPath}");
            return defaultPath;
        }

        return _settings.OdaExePath;
    }

    private sealed record CompareResult(List<DwgScanTarget> Pending, int Skipped);
}
