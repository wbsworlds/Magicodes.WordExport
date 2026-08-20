using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Magicodes.WordExport;
using Magicodes.WordExport.Models;
using Magicodes.WordExport.Print;
using ScottPlot;
using Image = System.Windows.Controls.Image;
using Label = System.Windows.Controls.Label;
using Orientation = System.Windows.Controls.Orientation;
using VerticalAlignment = System.Windows.VerticalAlignment;
using HorizontalAlignment = System.Windows.HorizontalAlignment;

namespace Magicodes.WordExport.Demo;

/// <summary>
/// WPF 主窗体：DataGrid 编辑样本 + ScottPlot 图表预览 + 生成/导出/打印预览按钮。
/// </summary>
public sealed class MainWindow : Window
{
    private readonly ObservableCollection<ReportDataFactory.Sample> _samples = new();
    private readonly DataGrid _grid = new();
    private readonly Image _chart = new();
    private readonly ListBox _log = new();

    private readonly string _templateDir = Path.Combine(AppContext.BaseDirectory, "Resources", "Templates");
    private readonly string _outputDir = Path.Combine(AppContext.BaseDirectory, "Outputs");
    private Plot _plot = new();

    public MainWindow()
    {
        Title = "Magicodes.WordExport Demo —— 检测报告生成器";
        Width = 1080;
        Height = 740;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        Directory.CreateDirectory(_outputDir);

        var root = new Grid { Margin = new Thickness(12) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(160) });

        root.Children.Add(BuildToolbar());
        Grid.SetRow(root.Children[^1], 0);

        var split = new GridSplitter { Width = 8, ResizeDirection = GridResizeDirection.Columns, Background = Brushes.Transparent, VerticalAlignment = VerticalAlignment.Stretch, HorizontalAlignment = HorizontalAlignment.Center };
        var splitGrid = new Grid();
        splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var gridPanel = new StackPanel();
        gridPanel.Children.Add(new Label { Content = "检测样本数据（编辑后图表自动刷新）", FontWeight = FontWeights.SemiBold });
        ConfigureGrid();
        gridPanel.Children.Add(_grid);
        splitGrid.Children.Add(gridPanel);
        Grid.SetColumn(gridPanel, 0);
        splitGrid.Children.Add(split);
        Grid.SetColumn(split, 1);
        var chartPanel = new StackPanel();
        chartPanel.Children.Add(new Label { Content = "ScottPlot 图表预览（实时）", FontWeight = FontWeights.SemiBold });
        _chart.Stretch = Stretch.Uniform;
        var chartBorder = new Border { Background = Brushes.White, Child = _chart };
        chartPanel.Children.Add(chartBorder);
        splitGrid.Children.Add(chartPanel);
        Grid.SetColumn(chartPanel, 2);

        root.Children.Add(splitGrid);
        Grid.SetRow(splitGrid, 1);

        var logPanel = new StackPanel();
        logPanel.Children.Add(new Label { Content = "操作日志", FontWeight = FontWeights.SemiBold });
        _log.FontFamily = new FontFamily("Consolas");
        _log.FontSize = 12;
        logPanel.Children.Add(_log);
        root.Children.Add(logPanel);
        Grid.SetRow(logPanel, 2);

        Content = root;

        Loaded += OnLoaded;
    }

    private UIElement BuildToolbar()
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 12) };
        void AddBtn(string text, RoutedEventHandler onClick)
        {
            var b = new Button { Content = text, Padding = new Thickness(14, 6, 14, 6), Margin = new Thickness(0, 0, 8, 0) };
            b.Click += onClick;
            panel.Children.Add(b);
        }
        AddBtn("生成 Word", (_, _) => GenerateWord());
        AddBtn("导出 PDF", (_, _) => ExportPdf());
        AddBtn("打印预览", (_, _) => Preview());
        AddBtn("打印", (_, _) => PrintDirect());
        AddBtn("打开模板目录", (_, _) => OpenTemplateDir());
        return panel;
    }

    private void ConfigureGrid()
    {
        _grid.AutoGenerateColumns = true;
        _grid.CanUserAddRows = false;
        _grid.CanUserDeleteRows = false;
        _grid.HeadersVisibility = DataGridHeadersVisibility.Column;
        _grid.CellEditEnding += (_, _) => Dispatcher.BeginInvoke(new Action(RefreshChart));
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        var tpl = TemplateBuilder.EnsureTemplate(_templateDir);
        Log($"✓ 模板已就绪: {tpl}");

        var (_, plot, samples) = ReportDataFactory.Create();
        _plot = plot;
        foreach (var s in samples) _samples.Add(s);
        _grid.ItemsSource = _samples;

        RefreshChart();
        Log($"✓ 图表预览已渲染");
    }

    private void RefreshChart()
    {
        var conc = _samples.Select(s => s.Concentration).ToArray();
        var abs = _samples.Select(s => s.Abs).ToArray();
        _plot = new Plot();
        if (conc.Length > 0) _plot.Add.Scatter(conc, abs);
        _plot.Title("标准曲线");
        _plot.XLabel("浓度 (mg/L)");
        _plot.YLabel("Abs");
        if (conc.Length > 0) _plot.Axes.SetLimitsX(0, Math.Max(3, conc.Max() * 1.1));
        if (abs.Length > 0) _plot.Axes.SetLimitsY(0, Math.Max(2.6, abs.Max() * 1.1));

        // 应用微软雅黑，避免中文标签乱码；标题/轴标签加粗
        _plot.ApplyCjkFont(bold: true);

        using var img = _plot.GetImage(720, 380);
        var png = img.GetImageBytes(ScottPlot.ImageFormat.Png, 100);
        var bi = new BitmapImage();
        bi.BeginInit();
        bi.CacheOption = BitmapCacheOption.OnLoad;
        bi.StreamSource = new MemoryStream(png);
        bi.EndInit();
        bi.Freeze();
        _chart.Source = bi;
    }

    private ReportData BuildReportData()
    {
        return new ReportData()
            .AddVariable("ReportNo", "RPT-2026-0815-001")
            .AddVariable("SampleName", "果蔬汁 - 维生素C")
            .AddVariable("TestType", "紫外分光光度法")
            .AddVariable("Project", "果蔬检测")
            .AddVariable("Unit", "mg/L")
            .AddVariable("Conclusion", "合格")
            .AddVariable("Formula", "y = 0.9997x + 0.0007")
            .AddVariable("Wavelength", "265")
            .AddVariable("PathLength", "1.0")
            .AddVariable("Tester", "张三")
            .AddVariable("Reviewer", "李四")
            .AddVariable("TestDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            .AddTable("Samples", _samples.Select(s => new System.Collections.Generic.Dictionary<string, object?>
            {
                ["Index"] = s.Index,
                ["Name"] = s.Name,
                ["Abs"] = s.Abs.ToString("F4"),
                ["Concentration"] = s.Concentration.ToString("F4"),
                ["Decision"] = s.Decision.ToString("F2"),
                ["Result"] = s.Result,
            }))
            .AddChart("1", _plot);
    }

    private WordExportBuilder BuildBuilder()
    {
        var tpl = TemplateBuilder.EnsureTemplate(_templateDir);
        return WordExportBuilder
            .Create()
            .UseTemplate(tpl)
            .Configure(o =>
            {
                //o.Orientation = PageOrientation.Landscape;
                o.PaperSize = PaperSize.A4;
                o.Margins = PageMargins.Min;
            })
            .WithData(BuildReportData());
    }

    private void GenerateWord()
    {
        try
        {
            var sw = Stopwatch.StartNew();
            var path = Path.Combine(_outputDir, $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.docx");
            BuildBuilder().SaveWord(path);
            Log(Diag.DumpDocxPageInfo(path));
            sw.Stop();
            Log($"✓ Word 已生成 ({sw.ElapsedMilliseconds} ms): {Path.GetFileName(path)}");
        }
        catch (Exception ex) { Log($"✗ 生成 Word 失败：{ex.Message}"); }
    }

    private void ExportPdf()
    {
        try
        {
            var sw = Stopwatch.StartNew();
            var path = Path.Combine(_outputDir, $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
            BuildBuilder().ExportPdf(path);
            sw.Stop();
            Log($"✓ PDF 已导出 ({sw.ElapsedMilliseconds} ms): {Path.GetFileName(path)}");
        }
        catch (Exception ex) { Log($"✗ 导出 PDF 失败：{ex.Message}"); }
    }

    private async void Preview()
    {
        try
        {
            var sw = Stopwatch.StartNew();
            var result = await BuildBuilder().PreviewAsync("检测报告");
            sw.Stop();
            Log(result.IsSuccess
                ? $"✓ 打印预览已关闭（共 {result.PaperCount} 页, {sw.ElapsedMilliseconds} ms）"
                : $"✗ 打印预览取消（{sw.ElapsedMilliseconds} ms）");
        }
        catch (Exception ex) { Log($"✗ 打印预览失败：{ex.Message}"); }
    }

    private async void PrintDirect()
    {
        try
        {
            var sw = Stopwatch.StartNew();
            var result = await BuildBuilder().PrintDirectAsync("检测报告", this);
            sw.Stop();
            Log(result.IsSuccess
                ? $"✓ 打印完成：{result.PaperCount} 页 ({sw.ElapsedMilliseconds} ms)"
                : $"✗ 打印取消（{sw.ElapsedMilliseconds} ms）");
        }
        catch (Exception ex) { Log($"✗ 打印失败：{ex.Message}"); }
    }

    private void OpenTemplateDir()
    {
        try { Process.Start(new ProcessStartInfo("explorer.exe", _templateDir) { UseShellExecute = true }); }
        catch (Exception ex) { Log($"✗ 打开目录失败：{ex.Message}"); }
    }

    private void Log(string msg)
        => _log.Items.Insert(0, $"[{DateTime.Now:HH:mm:ss}] {msg}");
}
