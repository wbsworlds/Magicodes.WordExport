using System;
using System.Collections.Generic;
using ScottPlot;

namespace Magicodes.WordExport.Models;

/// <summary>
/// Word 导出过程中使用的数据容器。
/// 支持变量、表格、图片（含 ScottPlot 图表）三种占位符。
/// </summary>
public sealed class ReportData
{
    /// <summary>变量字典。模板中 {{Key}} 会被替换。</summary>
    public Dictionary<string, object?> Variables { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>表格字典。Key 为表格名，模板中 {{Table>>Key|RowNo}}...{{Field|>>Table}} 配对出现。</summary>
    public Dictionary<string, IReadOnlyList<IDictionary<string, object?>>> Tables { get; } =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>图片字典。Key 为占位符编号，模板中 @image#Key:xxx.png 被替换。</summary>
    public Dictionary<string, ImageSource> Images { get; } = new(StringComparer.OrdinalIgnoreCase);

    public ReportData AddVariable(string key, object? value)
    {
        Variables[key] = value;
        return this;
    }

    public ReportData AddTable(string name, IEnumerable<IDictionary<string, object?>> rows)
    {
        Tables[name] = new List<IDictionary<string, object?>>(rows);
        return this;
    }

    public ReportData AddImage(string key, string filePath)
    {
        if (!System.IO.File.Exists(filePath))
            throw new System.IO.FileNotFoundException($"图片不存在: {filePath}", filePath);
        Images[key] = ImageSource.FromFile(filePath);
        return this;
    }

    public ReportData AddImage(string key, byte[] pngBytes)
    {
        Images[key] = ImageSource.FromBytes(pngBytes);
        return this;
    }

    /// <summary>添加 ScottPlot 图表。配置 lambda 在调用时执行。</summary>
    public ReportData AddChart(string key, Action<ScottPlot.Plot> configure, int width = 900, int height = 520)
    {
        var plot = new ScottPlot.Plot();
        configure(plot);
        Images[key] = ImageSource.FromChart(plot, width, height);
        return this;
    }

    /// <summary>添加一个已渲染好的 ScottPlot 图表（直接传入 Plot 对象）。</summary>
    public ReportData AddChart(string key, ScottPlot.Plot plot, int width = 900, int height = 520)
    {
        Images[key] = ImageSource.FromChart(plot, width, height);
        return this;
    }
}

/// <summary>
/// 图片来源：支持文件路径、字节、ScottPlot 图表三种。
/// </summary>
public sealed class ImageSource
{
    public string ContentType { get; }
    public byte[] Bytes { get; }
    public int PixelWidth { get; }
    public int PixelHeight { get; }
    public string SuggestedFileName { get; }

    private ImageSource(byte[] bytes, string contentType, int width, int height, string suggestedFileName)
    {
        Bytes = bytes;
        ContentType = contentType;
        PixelWidth = width;
        PixelHeight = height;
        SuggestedFileName = suggestedFileName;
    }

    public static ImageSource FromFile(string path)
    {
        var bytes = System.IO.File.ReadAllBytes(path);
        var ext = System.IO.Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        var ct = ext switch
        {
            "png" => "image/png",
            "jpg" or "jpeg" => "image/jpeg",
            "bmp" => "image/bmp",
            "gif" => "image/gif",
            _ => "application/octet-stream",
        };
        using var img = System.Drawing.Image.FromStream(new System.IO.MemoryStream(bytes));
        return new ImageSource(bytes, ct, img.Width, img.Height, System.IO.Path.GetFileName(path));
    }

    public static ImageSource FromBytes(byte[] bytes)
    {
        using var img = System.Drawing.Image.FromStream(new System.IO.MemoryStream(bytes));
        return new ImageSource(bytes, "image/png", img.Width, img.Height, "image.png");
    }

    public static ImageSource FromChart(ScottPlot.Plot plot, int width, int height)
    {
        // 渲染前统一应用中文字体，避免导出/打印的图表中文标签乱码。
        // 仅设置字体名，不改动调用方已设置的加粗等样式。
        plot.ApplyCjkFont();

        // ScottPlot 5.x 的 Plot.GetImage(w, h) 返回 ScottPlot.Image（IDisposable）。
        // 直接用 GetImageBytes(ImageFormat.Png, quality) 拿 PNG 字节。
        using var img = plot.GetImage(width, height);
        var bytes = img.GetImageBytes(ScottPlot.ImageFormat.Png, 100);
        return new ImageSource(bytes, "image/png", width, height, $"chart_{Guid.NewGuid():N}.png");
    }
}