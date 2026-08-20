using System;

namespace Magicodes.WordExport.Models;

/// <summary>
/// 全局导出配置。
///
/// 纸张 / 方向 / 页边距策略：
/// - 默认跟随 Word 模板（UseTemplateSettings=true）：
///   模板 docx 里原来的 PageSize / Orientation / Margins 保持不变。
/// - 调用者显式设置 PaperSize / Orientation / Margins 后：
///   对应字段生效，覆盖模板里的设置（逐项覆盖，未设置的仍用模板的）。
/// </summary>
public sealed class WordExportOptions
{
    /// <summary>Word 模板文件路径（.docx）。</summary>
    public string? TemplatePath { get; set; }

    /// <summary>是否在每页页脚添加页码。</summary>
    public bool AddPageNumber { get; set; } = true;

    /// <summary>
    /// 是否优先使用 Word 模板自身的纸张 / 方向 / 页边距设置（默认 true）。
    /// 若为 true：只有用户显式赋值了下面的 PaperSize / Orientation / Margins 才会覆盖模板对应项。
    /// 若为 false：模板里的页面设置全部忽略，所有项使用下面的值。
    /// </summary>
    public bool UseTemplateSettings { get; set; } = true;

    /// <summary>打印纸张：A4 / Letter / A3 等。null 表示跟随模板。</summary>
    public PaperSize? PaperSize { get; set; }

    /// <summary>打印方向。null 表示跟随模板。</summary>
    public PageOrientation? Orientation { get; set; }

    /// <summary>页边距（厘米）。null 表示跟随模板。</summary>
    public PageMargins? Margins { get; set; }

    /// <summary>PDF 导出质量（仅 FreeSpire.Doc 支持）。</summary>
    public PdfQuality PdfQuality { get; set; } = PdfQuality.High;

    /// <summary>渲染过程出现的非致命警告（收集起来给调用方）。</summary>
    public System.Collections.Generic.List<string> Warnings { get; } = new();
}

public enum PaperSize
{
    A3,
    A4,
    A5,
    Letter,
    Legal,
}

public enum PageOrientation
{
    Portrait,
    Landscape,
}

public enum PdfQuality
{
    Standard,
    High,
}

public sealed class PageMargins
{
    public double TopCm { get; set; } = 2.0;
    public double BottomCm { get; set; } = 2.0;
    public double LeftCm { get; set; } = 2.0;
    public double RightCm { get; set; } = 2.0;

    public static PageMargins Default => new();

    public static PageMargins Min => new()
    {
        TopCm = 1.0, BottomCm = 1.0, LeftCm = 1.0, RightCm = 1.0,
    };
}