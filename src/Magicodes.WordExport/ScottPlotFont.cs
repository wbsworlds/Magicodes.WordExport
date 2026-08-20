using ScottPlot;

namespace Magicodes.WordExport;

/// <summary>
/// ScottPlot 中文字体扩展。
///
/// 解决 headless 渲染（Plot.GetImage，无 WinForms FormsPlot 控件）时，
/// 中文标签以默认字体（缺中文字形）渲染出现乱码/方框的问题。
///
/// 等价于你给的 WinForms 写法，但换成 ScottPlot 5 的 Plot 对象 API：
///   formsPlot.Plot.Axes.Bottom.Label.FontName  ->  plot.Axes.Bottom.Label.FontName
///   formsPlot.Plot.Axes.Left.Label.FontName     ->  plot.Axes.Left.Label.FontName
///   formsPlot.Plot.Axes.Title.Label.FontName    ->  plot.Axes.Title.Label.FontName
///   formsPlot.Plot.Legend.FontName              ->  plot.Legend.FontName
///   formsPlot.Plot.Font.Set("...")              ->  plot.Font.Set("...")   // FontStyler.Set(string)，无 FontWeight 重载
/// 全局加粗用 LabelStyle.Bold = true 代替（ScottPlot 5.0.55 的 FontStyler 没有 Set(name, FontWeight) 重载）。
/// </summary>
public static class ScottPlotFont
{
    /// <summary>默认中文字体。Windows 自带，含完整中文字形。</summary>
    public const string DefaultCjkFont = "Microsoft YaHei UI";

    /// <summary>
    /// 为图表应用中文字体，避免标签乱码。
    /// </summary>
    /// <param name="plot">目标图表。</param>
    /// <param name="fontName">字体名，默认 Microsoft YaHei UI。</param>
    /// <param name="bold">是否加粗轴标签/标题/图例（全局加粗的等价做法）。</param>
    /// <returns>同一个 plot，便于链式调用。</returns>
    public static Plot ApplyCjkFont(this Plot plot, string fontName = DefaultCjkFont, bool bold = false)
    {
        if (plot == null) return plot!;

        // 全局默认字体：覆盖坐标轴刻度、轴标签、标题、图例等所有未单独指定的文字。
        // 字体不存在时静默跳过（仍走下面的逐项设置）。
        try { plot.Font.Set(fontName); }
        catch { /* 字体不可用：忽略全局设置，逐项设置仍生效 */ }

        // X 轴标签
        plot.Axes.Bottom.Label.FontName = fontName;
        // Y 轴标签
        plot.Axes.Left.Label.FontName = fontName;
        // 标题
        plot.Axes.Title.Label.FontName = fontName;
        // 图例
        plot.Legend.FontName = fontName;

        if (bold)
        {
            plot.Axes.Bottom.Label.Bold = true;
            plot.Axes.Left.Label.Bold = true;
            plot.Axes.Title.Label.Bold = true;
        }

        return plot;
    }
}
