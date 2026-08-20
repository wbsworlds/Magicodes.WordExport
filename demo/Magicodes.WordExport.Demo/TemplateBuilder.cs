using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Magicodes.WordExport.Demo;

/// <summary>
/// 程序化生成 report_template.docx 的辅助类。
/// 模板包含：
///   1. 顶部基础信息表
///   2. 中间图表占位符 @image#1:curve.png
///   3. 详细数据表（{{#TABLE_START:Samples}} + 字段占位符）
/// </summary>
public static class TemplateBuilder
{
    public static string EnsureTemplate(string targetDir)
    {
        Directory.CreateDirectory(targetDir);
        var templatePath = Path.Combine(targetDir, "report_template.docx");

        if (File.Exists(templatePath)) return templatePath;

        using var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document();
        var body = main.Document.AppendChild(new Body());

        // 标题
        AddParagraph(body, "检测报告", bold: true, size: "32", alignment: JustificationValues.Center);
        AddParagraph(body, " ");

        // 基础信息表
        BuildBasicInfoTable(body);
        AddParagraph(body, " ");

        // 图表标题
        AddParagraph(body, "标准曲线", bold: true, size: "24");
        // 图表占位符（替换为图片）
        AddParagraph(body, "@image#1:curve.png");
        AddParagraph(body, " ");

        // 详细数据标题
        AddParagraph(body, "检测详细数据", bold: true, size: "24");
        // 详细数据表格（示例行带 {{#TABLE_START:Samples}} 标记）
        BuildDataTable(body);
        AddParagraph(body, " ");

        // 页脚结论
        AddParagraph(body, "结论：{{Conclusion}}", bold: true);
        AddParagraph(body, "项目：{{Project}}    检测员：{{Tester}}    日期：{{TestDate}}", size: "20");

        // 页面设置：A4 横向
        var sectionProps = body.AppendChild(new SectionProperties());
        sectionProps.AppendChild(new PageSize
        {
            Width = 16838U,    // 横向 A4
            Height = 11906U,
            Orient = PageOrientationValues.Landscape,
        });
        sectionProps.AppendChild(new PageMargin
        {
            Top = 1000, Bottom = 1000, Left = 1200, Right = 1200,
            Header = 720, Footer = 720, Gutter = 0U,
        });

        main.Document.Save();
        return templatePath;
    }

    private static void BuildBasicInfoTable(Body body)
    {
        var table = body.AppendChild(new Table());
        var props = table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new LeftBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new RightBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" }
            ),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
        ));

        // 4 行 × 6 列的表头，每行单元格：label | value | label | value | label | value
        var rows = new[]
        {
            new[] { "编号", "{{ReportNo}}", "样品名", "{{SampleName}}", "检测类型", "{{TestType}}" },
            new[] { "项目", "{{Project}}", "单位", "{{Unit}}", "结论", "{{Conclusion}}" },
            new[] { "计算公式", "{{Formula}}", "波长", "{{Wavelength}} nm", "光程", "{{PathLength}} cm" },
            new[] { "检测员", "{{Tester}}", "复核员", "{{Reviewer}}", "检测时间", "{{TestDate}}" },
        };
        foreach (var r in rows)
        {
            var row = table.AppendChild(new TableRow());
            for (int i = 0; i < r.Length; i++)
            {
                var cell = row.AppendChild(new TableCell());
                cell.AppendChild(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = "0" },
                    new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = i % 2 == 0 ? "EFEFEF" : "FFFFFF" }
                ));
                var p = cell.AppendChild(new Paragraph());
                if (i % 2 == 0)
                {
                    p.AppendChild(new Run(new RunProperties(new Bold()), new Text(r[i])));
                }
                else
                {
                    p.AppendChild(new Run(new Text(r[i])));
                }
            }
        }
    }

    private static void BuildDataTable(Body body)
    {
        var table = body.AppendChild(new Table());
        table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new LeftBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new RightBorder { Val = BorderValues.Single, Size = 6, Color = "888888" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" }
            ),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
        ));

        // 表头行
        var headers = new[] { "序号", "品名", "Abs", "浓度", "决策值", "结果" };
        var headerRow = table.AppendChild(new TableRow());
        headerRow.AppendChild(new TableRowProperties(new TableHeader()));
        foreach (var h in headers)
        {
            var cell = headerRow.AppendChild(new TableCell());
            cell.AppendChild(new TableCellProperties(
                new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "DDE6F1" }
            ));
            cell.AppendChild(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(new RunProperties(new Bold()), new Text(h))
            ));
        }

        // 示例数据行（含 {{#TABLE_START:Samples}} 标记 + {{Index}} 字段）
        // 标记单独一段，{{Index}} 单独一段 —— ProcessTables 只清标记段，不会误伤字段段
        var dataRow = table.AppendChild(new TableRow());
        var indexCell = new TableCell();
        indexCell.AppendChild(new Paragraph(new Run(new Text("{{#TABLE_START:Samples}}"))));
        indexCell.AppendChild(new Paragraph(new Run(new Text("{{Index}}"))));
        dataRow.AppendChild(indexCell);
        dataRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("{{Name}}")))));
        dataRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("{{Abs}}")))));
        dataRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("{{Concentration}}")))));
        dataRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("{{Decision}}")))));
        dataRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("{{Result}}")))));
    }

    private static void AddParagraph(Body body, string text, bool bold = false, string size = "22")
        => AddParagraph(body, text, bold, size, null);

    private static void AddParagraph(Body body, string text, bool bold, string size, JustificationValues? alignment)
    {
        var p = body.AppendChild(new Paragraph());
        if (alignment.HasValue)
            p.AppendChild(new ParagraphProperties(new Justification { Val = alignment.Value }));
        var run = p.AppendChild(new Run());
        if (bold || size != "22")
        {
            var rp = run.AppendChild(new RunProperties());
            if (bold) rp.AppendChild(new Bold());
            if (size != "22") rp.AppendChild(new FontSize { Val = size });
        }
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }
}