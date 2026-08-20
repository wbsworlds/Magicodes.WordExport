using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Printing;
using WpfMessageBox = System.Windows.MessageBox;
using Image = System.Drawing.Image;
using Size = System.Windows.Size;

namespace Magicodes.WordExport.Print;

/// <summary>
/// 打印预览窗口，使用 WPF 原生 DocumentViewer 控件。
/// 内置打印、缩放、页面导航、缩略图，外观专业，无第三方依赖。
/// 每页尺寸根据 PreviewPageInfo.PhysicalSize 动态计算，预览与实际文档一致。
/// 打印按钮弹出自绘设置窗口（FluentPrintSettingsWindow），不使用 Win32 通用对话框，
/// 因此不会有"此应用不支持打印预览"的提示。
/// </summary>
public sealed class FluentPrintPreviewWindow : Window
{
    private readonly List<RenderedPage> _pages;
    private readonly string _documentName;

    /// <summary>打印结果。</summary>
    public FluentPrintResult Result { get; private set; } = new(false, 0);

    /// <param name="pages">已渲染的分页数据（含 GDI 图像 + WPF 页面尺寸）。</param>
    /// <param name="documentName">文档名称。</param>
    public FluentPrintPreviewWindow(List<RenderedPage> pages, string documentName)
    {
        _pages = pages;
        _documentName = documentName;

        Title = $"打印预览 —— {documentName}";
        Width = 1024;
        Height = 700;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        WindowState = WindowState.Maximized;
        ResizeMode = ResizeMode.CanResize;

        var viewer = new DocumentViewer { Document = BuildFixedDocument() };

        // 覆盖 DocumentViewer 默认的 Win32 打印命令，改用自绘设置窗口
        viewer.CommandBindings.Add(new CommandBinding(
            ApplicationCommands.Print,
            (_, _) => ShowCustomPrintDialog(viewer)));

        Closed += (_, _) => { Result = new FluentPrintResult(true, _pages.Count); };
        Content = viewer;
    }

    /// <summary>
    /// 弹出自绘打印设置窗口，然后用 WPF PrintDialog（只取 PrintQueue/PrintTicket）发送到打印机。
    /// </summary>
    private void ShowCustomPrintDialog(DocumentViewer viewer)
    {
        if (_pages.Count == 0) return;

        var firstPage = _pages[0];
        var defaultOrientation = firstPage.PageWidthWpf > firstPage.PageHeightWpf
            ? PageOrientation.Landscape
            : PageOrientation.Portrait;

        var settingsWin = new FluentPrintSettingsWindow(_pages.Count, defaultOrientation)
        {
            Owner = this,
        };
        if (settingsWin.ShowDialog() != true || settingsWin.Settings == null) return;

        var s = settingsWin.Settings;
        try
        {
            PrintPagesWpf(s, _pages, _documentName);
        }
        catch (Exception ex)
        {
            WpfMessageBox.Show(this, $"打印失败：{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 使用 PrintDialog.PrintDocument 方式打印，正确传递方向/纸张尺寸。
    /// 不传系统 PrintDialog 的 UI，只借其 API 投递作业。
    /// </summary>
    internal static void PrintPagesWpf(PrintSettings s, List<RenderedPage> pages, string documentName)
    {
        int fromIdx = s.FromPage - 1;
        int toIdx = s.ToPage - 1;

        var subset = new List<RenderedPage>(capacity: s.Copies * (toIdx - fromIdx + 1));
        for (int c = 0; c < s.Copies; c++)
            for (int i = fromIdx; i <= toIdx; i++)
                subset.Add(pages[i]);

        // 构建临时 FixedDocument
        var doc = new FixedDocument();
        if (subset.Count == 0) return;
        var first = subset[0];
        doc.DocumentPaginator.PageSize = new Size(first.PageWidthWpf, first.PageHeightWpf);
        foreach (var rp in subset)
        {
            var bi = GdiToBitmapSource(rp.Image);
            var fp = new FixedPage { Width = rp.PageWidthWpf, Height = rp.PageHeightWpf };
            var img = new System.Windows.Controls.Image
            {
                Source = bi,
                Stretch = Stretch.Fill,
                Width = fp.Width,
                Height = fp.Height,
            };
            RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);
            FixedPage.SetLeft(img, 0); FixedPage.SetTop(img, 0);
            fp.Children.Add(img);
            var pc = new PageContent { Child = fp };
            doc.Pages.Add(pc);
        }

        // 只借 PrintDialog API：指定打印机和方向，不弹窗（不调用 ShowDialog）
        var printDlg = new System.Windows.Controls.PrintDialog
        {
            PrintQueue = s.PrintQueue,
        };
        printDlg.PrintTicket = new PrintTicket
        {
            PageOrientation = s.Orientation,
            CopyCount = s.Copies,
        };
        printDlg.PrintDocument(doc.DocumentPaginator, documentName);
    }

    /// <summary>
    /// 将 GDI 分页图像构建为 FixedDocument，每页使用文档实际的页面尺寸。
    /// </summary>
    private FixedDocument BuildFixedDocument()
    {
        var doc = new FixedDocument();
        if (_pages.Count == 0) return doc;

        var firstPage = _pages[0];
        doc.DocumentPaginator.PageSize = new Size(firstPage.PageWidthWpf, firstPage.PageHeightWpf);

        foreach (var rp in _pages)
        {
            var page = BuildFixedPage(rp);
            var pageContent = new PageContent { Child = page };
            doc.Pages.Add(pageContent);
        }
        return doc;
    }

    private FixedPage BuildFixedPage(RenderedPage rp)
    {
        var bi = GdiToBitmapSource(rp.Image);
        var page = new FixedPage
        {
            Width = rp.PageWidthWpf,
            Height = rp.PageHeightWpf,
        };

        var img = new System.Windows.Controls.Image
        {
            Source = bi,
            Stretch = Stretch.Fill,
        };
        RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);

        FixedPage.SetLeft(img, 0);
        FixedPage.SetTop(img, 0);
        img.Width = page.Width;
        img.Height = page.Height;

        page.Children.Add(img);
        return page;
    }

    /// <summary>GDI Image → BitmapSource（冻结，跨线程安全）。</summary>
    internal static BitmapSource GdiToBitmapSource(Image gdi)
    {
        using var ms = new MemoryStream();
        gdi.Save(ms, ImageFormat.Png);
        ms.Position = 0;

        var bi = new BitmapImage();
        bi.BeginInit();
        bi.CacheOption = BitmapCacheOption.OnLoad;
        bi.StreamSource = ms;
        bi.EndInit();
        bi.Freeze();
        return bi;
    }
}