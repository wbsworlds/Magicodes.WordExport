using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Printing;
using WpfMessageBox = System.Windows.MessageBox;
using Image = System.Drawing.Image;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using ProgressBar = System.Windows.Controls.ProgressBar;

namespace Magicodes.WordExport.Print;

/// <summary>打印结果。</summary>
public sealed record FluentPrintResult(bool IsSuccess, int PaperCount);

/// <summary>
/// 已渲染的单页数据：GDI 图像 + WPF 页面尺寸（由 PreviewPageInfo.PhysicalSize 转换而来）。
/// </summary>
public sealed record RenderedPage(Image Image, double PageWidthWpf, double PageHeightWpf);

/// <summary>
/// Fluent 风格打印入口。
/// - 预览：DocumentViewer + 自绘打印设置窗口（无 Win32"不支持预览"）
/// - 直接打印：渲染后直接弹自绘设置窗口 → 打印
/// 不依赖 PrintDialogX / Wpf.Ui，无第三方库绑定错误。
/// </summary>
public sealed class FluentPrintDialog
{
    private readonly DocumentPrinter _documentPrinter = new();

    /// <summary>打开打印预览对话框（模态）。</summary>
    public FluentPrintResult ShowDialog(byte[] docxBytes, Models.WordExportOptions options, string documentName = "Word 文档")
    {
        return ShowDialogAsync(docxBytes, options, documentName).GetAwaiter().GetResult();
    }

    /// <summary>
    /// 异步打开打印预览对话框。
    /// GDI 渲染在后台线程执行，UI 显示加载提示，完成后打开自定义预览窗口。
    /// </summary>
    public async Task<FluentPrintResult> ShowDialogAsync(byte[] docxBytes, Models.WordExportOptions options, string documentName = "Word 文档")
    {
        var loadingWindow = CreateLoadingWindow();
        loadingWindow.Show();

        List<RenderedPage> pages;
        try
        {
            pages = await Task.Run(() => RenderGdiPages(docxBytes));
        }
        finally
        {
            loadingWindow.Close();
        }

        if (pages.Count == 0) return new FluentPrintResult(false, 0);

        var previewWindow = new FluentPrintPreviewWindow(pages, documentName);
        previewWindow.ShowDialog();
        return previewWindow.Result;
    }

    /// <summary>
    /// 直接打印（不打开预览）。
    /// GDI 渲染 → 加载提示 → 自绘打印设置窗口 → 打印到所选打印机。
    /// </summary>
    public FluentPrintResult PrintDirect(byte[] docxBytes, Models.WordExportOptions options, string documentName = "Word 文档", Window? owner = null)
    {
        return PrintDirectAsync(docxBytes, options, documentName, owner).GetAwaiter().GetResult();
    }

    /// <summary>异步直接打印。</summary>
    public async Task<FluentPrintResult> PrintDirectAsync(byte[] docxBytes, Models.WordExportOptions options, string documentName = "Word 文档", Window? owner = null)
    {
        var loadingWindow = CreateLoadingWindow();
        if (owner != null) loadingWindow.Owner = owner;
        loadingWindow.Show();

        List<RenderedPage> pages;
        try
        {
            pages = await Task.Run(() => RenderGdiPages(docxBytes));
        }
        finally
        {
            loadingWindow.Close();
        }

        if (pages.Count == 0) return new FluentPrintResult(false, 0);

        var first = pages[0];
        var defaultOrient = first.PageWidthWpf > first.PageHeightWpf
            ? PageOrientation.Landscape
            : PageOrientation.Portrait;

        var settingsWin = new FluentPrintSettingsWindow(pages.Count, defaultOrient);
        if (owner != null) settingsWin.Owner = owner;
        if (settingsWin.ShowDialog() != true || settingsWin.Settings == null)
            return new FluentPrintResult(false, pages.Count);

        try
        {
            FluentPrintPreviewWindow.PrintPagesWpf(settingsWin.Settings, pages, documentName);
            return new FluentPrintResult(true, pages.Count);
        }
        catch (Exception ex)
        {
            WpfMessageBox.Show(owner, $"打印失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return new FluentPrintResult(false, pages.Count);
        }
    }

    /// <summary>创建加载提示窗口。</summary>
    private static Window CreateLoadingWindow()
    {
        var panel = new StackPanel
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };
        panel.Children.Add(new TextBlock
        {
            Text = "正在渲染文档，请稍候…",
            FontSize = 16,
            Margin = new Thickness(0, 0, 0, 16),
            HorizontalAlignment = HorizontalAlignment.Center,
        });
        var pb = new ProgressBar
        {
            IsIndeterminate = true,
            Width = 240,
            Height = 4,
        };
        panel.Children.Add(pb);

        return new Window
        {
            Title = "渲染中",
            Content = panel,
            Width = 340,
            Height = 140,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            Topmost = true,
        };
    }

    /// <summary>
    /// docx → 分页 GDI 图像列表 + WPF 页面尺寸。
    /// PreviewPageInfo.PhysicalSize 是 1/100 英寸单位，需转为 WPF 的 1/96 英寸单位。
    /// </summary>
    private List<RenderedPage> RenderGdiPages(byte[] docxBytes)
    {
        var result = new List<RenderedPage>();
        using var pd = _documentPrinter.CreatePrintDocument(docxBytes);
        var preview = new PreviewPrintController { UseAntiAlias = true };
        pd.PrintController = preview;
        pd.Print();

        foreach (var info in preview.GetPreviewPageInfo())
        {
            if (info.Image == null) continue;
            double wpfWidth = info.PhysicalSize.Width / 100.0 * 96.0;
            double wpfHeight = info.PhysicalSize.Height / 100.0 * 96.0;
            result.Add(new RenderedPage(info.Image, wpfWidth, wpfHeight));
        }
        return result;
    }
}