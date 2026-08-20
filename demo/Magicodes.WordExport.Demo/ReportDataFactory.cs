using System;
using System.Collections.Generic;
using Magicodes.WordExport;
using Magicodes.WordExport.Models;
using ScottPlot;

namespace Magicodes.WordExport.Demo;

/// <summary>
/// 构造演示用的检测报告数据。模拟截图里的"吸光度 vs 浓度"标准曲线场景。
/// </summary>
public static class ReportDataFactory
{
    public sealed class Sample
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public double Abs { get; set; }
        public double Concentration { get; set; }
        public double Decision { get; set; }
        public string Result { get; set; } = "";
    }

    /// <summary>构造报告数据 + ScottPlot 曲线 plot。</summary>
    public static (ReportData data, Plot plot, IReadOnlyList<Sample> samples) Create()
    {
        var samples = new List<Sample>
        {
            new() { Index = 1, Name = "样品1",  Abs = 0.0001, Concentration = 0.0001, Decision = 1.0, Result = "合格" },
            new() { Index = 2, Name = "样品2",  Abs = 0.0155, Concentration = 0.0155, Decision = 1.0, Result = "合格" },
            new() { Index = 3, Name = "样品3",  Abs = 0.5001, Concentration = 0.5001, Decision = 1.0, Result = "合格" },
            new() { Index = 4, Name = "样品4",  Abs = 1.0001, Concentration = 1.0001, Decision = 1.0, Result = "合格" },
            new() { Index = 5, Name = "样品5",  Abs = 1.5001, Concentration = 1.5001, Decision = 1.0, Result = "合格" },
            new() { Index = 6, Name = "样品6",  Abs = 1.9991, Concentration = 1.9991, Decision = 1.0, Result = "合格" },
            new() { Index = 7, Name = "样品7",  Abs = 2.5001, Concentration = 2.5001, Decision = 1.0, Result = "合格" },
        };

        // 拟合线 y = x
        var conc = samples.ConvertAll(s => s.Concentration);
        var abs = samples.ConvertAll(s => s.Abs);

        // ScottPlot 图表
        var plot = new Plot();
        plot.Add.Scatter(conc, abs);
        plot.Title("标准曲线");
        plot.XLabel("浓度 (mg/L)");
        plot.YLabel("Abs");
        plot.Axes.SetLimitsX(0, 3);
        plot.Axes.SetLimitsY(0, 2.6);

        // 应用微软雅黑，避免中文标签乱码；标题/轴标签加粗
        plot.ApplyCjkFont(bold: true);

        // 用类库 API 构造 ReportData（直接构造 Model 避免 builder 状态被破坏）
        var data = new ReportData()
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
            .AddVariable("TestDate", "2026-08-15 12:00:00")
            .AddTable("Samples", samples.Select(s => new Dictionary<string, object?>
            {
                ["Index"] = s.Index,
                ["Name"] = s.Name,
                ["Abs"] = s.Abs.ToString("F4"),
                ["Concentration"] = s.Concentration.ToString("F4"),
                ["Decision"] = s.Decision.ToString("F2"),
                ["Result"] = s.Result,
            }))
            .AddChart("1", plot);

        return (data, plot, samples);
    }
}