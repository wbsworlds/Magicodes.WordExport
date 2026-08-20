using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Magicodes.WordExport.Models;
using Magicodes.WordExport.Pdf;
using Magicodes.WordExport.Print;
using Magicodes.WordExport.Render;

namespace Magicodes.WordExport;

/// <summary>
/// Word 导出流式 API 主入口。
///
/// 示例：
/// <code>
/// var result = WordExportBuilder
///     .Create()
///     .UseTemplate("template.docx")
///     .AddVariable("ReportNo", "RPT-001")
///     .AddTable("Samples", samples)
///     .AddImage("logo", "logo.png")
///     .AddChart("curve", p => { p.Add.Scatter(xs, ys); })
///     .Build();
/// result.SaveWordAs("report.docx");
/// </code>
/// </summary>
public sealed class WordExportBuilder
{
    private readonly ReportData _data = new();
    private readonly WordExportOptions _options = new();
    private byte[]? _templateBytes;
    private string? _templatePath;

    private WordExportBuilder() { }

    public static WordExportBuilder Create() => new();

    // ============================================================
    // 模板
    // ============================================================

    public WordExportBuilder UseTemplate(string templatePath)
    {
        if (string.IsNullOrWhiteSpace(templatePath))
            throw new ArgumentException("模板路径不能为空", nameof(templatePath));
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"模板文件不存在: {templatePath}", templatePath);
        _templatePath = templatePath;
        return this;
    }

    public WordExportBuilder UseTemplate(byte[] templateBytes)
    {
        _templateBytes = templateBytes ?? throw new ArgumentNullException(nameof(templateBytes));
        return this;
    }

    // ============================================================
    // 全局选项
    // ============================================================

    public WordExportBuilder Configure(Action<WordExportOptions> configure)
    {
        configure?.Invoke(_options);
        return this;
    }

    // ============================================================
    // 数据快捷方法
    // ============================================================

    public WordExportBuilder WithData(ReportData data)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));
        foreach (var kv in data.Variables) _data.Variables[kv.Key] = kv.Value;
        foreach (var kv in data.Tables) _data.Tables[kv.Key] = kv.Value;
        foreach (var kv in data.Images) _data.Images[kv.Key] = kv.Value;
        return this;
    }

    public WordExportBuilder AddVariable(string key, object? value)
    {
        _data.AddVariable(key, value);
        return this;
    }

    public WordExportBuilder AddTable(string name, IEnumerable<IDictionary<string, object?>> rows)
    {
        _data.AddTable(name, rows);
        return this;
    }

    public WordExportBuilder AddImage(string key, string filePath)
    {
        _data.AddImage(key, filePath);
        return this;
    }

    public WordExportBuilder AddImage(string key, byte[] pngBytes)
    {
        _data.AddImage(key, pngBytes);
        return this;
    }

    public WordExportBuilder AddChart(string key, Action<ScottPlot.Plot> configure, int width = 900, int height = 520)
    {
        _data.AddChart(key, configure, width, height);
        return this;
    }

    public WordExportBuilder AddChart(string key, ScottPlot.Plot plot, int width = 900, int height = 520)
    {
        _data.AddChart(key, plot, width, height);
        return this;
    }

    // ============================================================
    // 渲染入口
    // ============================================================

    /// <summary>同步渲染，返回 Word 字节。</summary>
    public RenderResult Build()
    {
        return BuildInternal();
    }

    /// <summary>异步渲染，返回 Word 字节。</summary>
    public Task<RenderResult> BuildAsync()
    {
        return Task.FromResult(BuildInternal());
    }

    private RenderResult BuildInternal()
    {
        var templateBytes = LoadTemplateBytes();
        var renderer = new TemplateRenderer(_data, _options);
        var wordBytes = renderer.Render(templateBytes);

        var outDir = Path.Combine(Path.GetTempPath(), "Magicodes.WordExport", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outDir);
        return new RenderResult(wordBytes, outDir);
    }

    private byte[] LoadTemplateBytes()
    {
        if (_templateBytes != null) return _templateBytes;
        if (_templatePath != null) return File.ReadAllBytes(_templatePath);
        throw new InvalidOperationException("未指定模板。请先调用 UseTemplate(...).");
    }

    // ============================================================
    // 输出快捷方法
    // ============================================================

    /// <summary>渲染并保存为 Word 文件。</summary>
    public string SaveWord(string path)
    {
        var r = Build();
        return r.SaveWordAs(path);
    }

    /// <summary>渲染并保存为 Word 文件（异步）。</summary>
    public Task<string> SaveWordAsync(string path)
    {
        return Task.FromResult(SaveWord(path));
    }

    /// <summary>渲染并保存为 PDF 文件。</summary>
    public string ExportPdf(string path)
    {
        var r = Build();
        var pdfBytes = new PdfExporter().ExportToPdfBytes(r.WordBytes, _options);
        r.PdfBytes = pdfBytes;
        File.WriteAllBytes(path, pdfBytes);
        return path;
    }

    /// <summary>渲染并保存为 PDF 文件（异步）。</summary>
    public Task<string> ExportPdfAsync(string path)
    {
        return Task.FromResult(ExportPdf(path));
    }

    /// <summary>
    /// 渲染并打开自定义打印预览窗口：分页预览 + 打印机选择 + 打印设置 + 打印一体。
    /// 需 STA 线程。旧的 WinForms PrintPreviewDialog 仍可用 DocumentPrinter.ShowPreview 访问。
    /// </summary>
    public FluentPrintResult Preview(string documentName = "Word 文档")
    {
        var r = Build();
        return new FluentPrintDialog().ShowDialog(r.WordBytes, _options, documentName);
    }

    /// <summary>
    /// 异步渲染并打开打印预览窗口：
    /// GDI 分页渲染在后台线程执行，UI 显示加载提示，完成后打开自定义预览窗口。
    /// </summary>
    public async Task<FluentPrintResult> PreviewAsync(string documentName = "Word 文档")
    {
        var r = Build();
        return await new FluentPrintDialog().ShowDialogAsync(r.WordBytes, _options, documentName);
    }

    /// <summary>渲染并使用默认打印机打印。</summary>
    public void Print()
    {
        var r = Build();
        new DocumentPrinter().Print(r.WordBytes, _options);
    }

    /// <summary>渲染并打印到指定打印机。</summary>
    public void Print(string printerName, int copies = 1)
    {
        var r = Build();
        new DocumentPrinter().PrintOn(r.WordBytes, _options, printerName, copies);
    }

    /// <summary>渲染并打印到指定打印机（异步）。</summary>
    public Task PrintAsync(string printerName, int copies = 1)
    {
        return Task.Run(() => Print(printerName, copies));
    }

    /// <summary>
    /// 直接打印（不预览）：渲染后弹出自绘打印设置窗口，由用户选择打印机/方向/份数/页码后打印。
    /// 没有 Win32 对话框的"此应用不支持打印预览"提示。
    /// </summary>
    public FluentPrintResult PrintDirect(string documentName = "Word 文档", System.Windows.Window? owner = null)
    {
        var r = Build();
        return new FluentPrintDialog().PrintDirect(r.WordBytes, _options, documentName, owner);
    }

    /// <summary>直接打印（异步）：GDI 渲染在后台线程，完成后弹自绘设置窗口。</summary>
    public async Task<FluentPrintResult> PrintDirectAsync(string documentName = "Word 文档", System.Windows.Window? owner = null)
    {
        var r = Build();
        return await new FluentPrintDialog().PrintDirectAsync(r.WordBytes, _options, documentName, owner);
    }
}