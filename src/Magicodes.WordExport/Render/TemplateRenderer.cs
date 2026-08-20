using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Magicodes.WordExport.Models;
using Magicodes.WordExport.Template;

namespace Magicodes.WordExport.Render;

/// <summary>
/// 核心模板渲染器。基于 DocumentFormat.OpenXml 直接处理 docx。
///
/// 处理流程：
///   1. 打开模板 docx
///   2. 遍历所有段落/表格元素
///   3. 替换文本中的 {{Var}} 与 @image#KEY:filename.png
///   4. 把 @image 占位符所在的段落替换为内嵌图片 Run
///   5. 找到含 {{#TABLE_START:Name}} 的表格行，复制 N-1 次后回填字段
///   6. 应用页面方向 / 纸张 / 页边距设置
///   7. 保存为新的 docx 字节流
/// </summary>
public sealed class TemplateRenderer
{
    private readonly ReportData _data;
    private readonly WordExportOptions _options;

    public TemplateRenderer(ReportData data, WordExportOptions options)
    {
        _data = data;
        _options = options;
    }

    public byte[] Render(byte[] templateBytes)
    {
        if (templateBytes == null || templateBytes.Length == 0)
            throw new ArgumentException("模板字节为空", nameof(templateBytes));

        using var input = new MemoryStream(templateBytes);
        using var output = new MemoryStream();
        input.CopyTo(output);
        output.Position = 0;

        using (var doc = WordprocessingDocument.Open(output, true))
        {
            var body = doc.MainDocumentPart!.Document.Body!;

            // 关键顺序：
            //   1) 表格处理先做：把 {{Field}} 在表格行内替换为具体数据值
            //   2) 段落处理后做：避免变量字典的 {{Var}} 误伤表格字段（即使字段名相同也不冲突）
            // 3) 页面设置
            ProcessTables(body);
            ProcessParagraphs(body, doc.MainDocumentPart!);
            ApplyPageSettings(body, _options);

            doc.MainDocumentPart!.Document.Save();
        }

        return output.ToArray();
    }

    // ============================================================
    // 段落处理：变量 / 图片占位符
    // ============================================================

    private void ProcessParagraphs(Body body, MainDocumentPart mainPart)
    {
        // 收集所有段落（包括表格内的）
        var paragraphs = body.Descendants<Paragraph>().ToList();

        // 先把所有变量 / 图片占位符的文本替换完成
        foreach (var p in paragraphs)
        {
            ReplaceRunText(p, _data);
        }

        // 再统一处理需要插入图片的段落（@image#KEY:filename 已经变成临时记号）
        var imageParaKeys = new List<(Paragraph Para, string Key)>();
        foreach (var p in paragraphs)
        {
            foreach (var run in p.Elements<Run>().ToList())
            {
                foreach (var text in run.Elements<Text>().ToList())
                {
                    if (TemplateSyntax.TryUnwrapImageSentinel(text.Text, out var key)
                        && _data.Images.ContainsKey(key))
                    {
                        imageParaKeys.Add((p, key));
                        break;
                    }
                }
            }
        }

        // 把含临时记号的 Run 替换为图片 Run
        foreach (var (para, key) in imageParaKeys.DistinctBy(x => x.Para))
        {
            InsertImageIntoParagraph(para, key, mainPart);
        }
    }

    /// <summary>
    /// 把段落内的变量 / 图片占位符替换为值 / 临时记号。
    ///
    /// 关键：Word/WPF 保存 docx 时经常把连续文本拆成多个 <w:r><w:t> 段，
    /// 例如 "@image#1:curve.png" 可能变成 "@ima" + "ge#1:curve" + ".png"。
    /// 必须先把段落内全部 Text 拼接成整体再做匹配；若发现占位符 / 变量命中，
    /// 就把整段 Run 合并成单个 Run，避免跨段正则不匹配。
    /// </summary>
    private static void ReplaceRunText(Paragraph paragraph, ReportData data)
    {
        // 收集所有 Text 元素及其所属 Run
        var texts = paragraph.Descendants<Text>().ToList();
        if (texts.Count == 0) return;

        // 先整体拼接成一段文本（模拟用户看到的内容）
        var combined = string.Concat(texts.Select(t => t.Text));
        if (string.IsNullOrEmpty(combined)) return;

        var preprocessed = TemplatePreprocessor.PreprocessRunText(combined, data);

        // 如果整体文本都没有任何替换命中（图片+变量都没替换），
        // 则按老逻辑逐段尝试（纯变量逐段命中也能生效）
        if (preprocessed.Text == combined)
        {
            foreach (var text in texts)
            {
                if (string.IsNullOrEmpty(text.Text)) continue;
                var one = TemplatePreprocessor.PreprocessRunText(text.Text, data);
                if (one.Text != text.Text)
                {
                    text.Text = one.Text;
                    text.Space = SpaceProcessingModeValues.Preserve;
                }
            }
            return;
        }

        // 整体命中：把段落内所有 Run 清空，替换成一个合并后的新 Run
        // 通过从第一个非空 Run 拷贝 RunProperties 保留原有的字体/颜色/大小格式
        RunProperties? firstProps = null;
        foreach (var run in paragraph.Elements<Run>())
        {
            if (firstProps == null && run.RunProperties != null)
                firstProps = (RunProperties)run.RunProperties.CloneNode(true);
            paragraph.RemoveChild(run);
        }

        var mergedRun = new Run();
        if (firstProps != null) mergedRun.RunProperties = firstProps;
        mergedRun.AppendChild(new Text(preprocessed.Text)
        {
            Space = SpaceProcessingModeValues.Preserve,
        });
        paragraph.AppendChild(mergedRun);
    }

    // ============================================================
    // 表格处理：{{#TABLE_START:Name}} 标记 + 行复制 + 字段替换
    // ============================================================

    private void ProcessTables(Body body)
    {
        foreach (var table in body.Descendants<Table>().ToList())
        {
            TableRow? templateRow = null;
            string? tableName = null;

            foreach (var row in table.Elements<TableRow>().ToList())
            {
                var markerText = string.Concat(row.Descendants<Text>().Select(t => t.Text));
                var match = System.Text.RegularExpressions.Regex.Match(
                    markerText, @"\{\{\s*#TABLE_START\s*:\s*(?<name>[A-Za-z0-9_\-]+)\s*\}\}");
                if (match.Success)
                {
                    templateRow = row;
                    tableName = match.Groups["name"].Value;
                    foreach (var t in row.Descendants<Text>().ToList())
                    {
                        t.Text = t.Text.Replace(match.Value, "").Trim();
                    }
                    break;
                }
            }

            if (templateRow == null || tableName == null) continue;
            if (!_data.Tables.TryGetValue(tableName, out var dataRows) || dataRows.Count == 0)
            {
                continue;
            }

            // 关键：先复制再填充！否则 CloneNode(true) 会把已替换的文本复制到后续行，
            // 后续行就找不到 {{Field}} 占位符了。
            // 1) 先把模板行复制到 N 行
            var rows = new List<TableRow> { templateRow };
            for (int i = 1; i < dataRows.Count; i++)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);
                table.AppendChild(newRow);
                rows.Add(newRow);
            }

            // 2) 逐行填充字段
            for (int i = 0; i < dataRows.Count && i < rows.Count; i++)
            {
                FillRowFields(rows[i], dataRows[i]);
            }
        }
    }

    /// <summary>
    /// 把表格行内的 {{FieldName}} 替换为对应数据。
    /// 同样处理 Run 拆分问题：先把单元格内所有 Text 拼接，若命中则合并 Run。
    /// </summary>
    private static void FillRowFields(TableRow row, IDictionary<string, object?> data)
    {
        // 对每个 Paragraph 做一次整体扫描处理（逐单元格逐段落）
        foreach (var para in row.Descendants<Paragraph>().ToList())
        {
            var texts = para.Descendants<Text>().ToList();
            if (texts.Count == 0) continue;

            var combined = string.Concat(texts.Select(t => t.Text));
            if (string.IsNullOrEmpty(combined)) continue;

            var replaced = System.Text.RegularExpressions.Regex.Replace(
                combined, @"\{\{\s*(?<name>[A-Za-z0-9_\.]+)\s*\}\}", m =>
                {
                    var name = m.Groups["name"].Value;
                    if (data.TryGetValue(name, out var v)) return v?.ToString() ?? "";
                    return m.Value;
                });

            if (replaced == combined)
            {
                // 没有整体命中，继续逐段替换
                foreach (var text in texts)
                {
                    if (string.IsNullOrEmpty(text.Text)) continue;
                    var one = System.Text.RegularExpressions.Regex.Replace(
                        text.Text, @"\{\{\s*(?<name>[A-Za-z0-9_\.]+)\s*\}\}", m =>
                        {
                            var name = m.Groups["name"].Value;
                            if (data.TryGetValue(name, out var v)) return v?.ToString() ?? "";
                            return m.Value;
                        });
                    if (one != text.Text)
                    {
                        text.Text = one;
                        text.Space = SpaceProcessingModeValues.Preserve;
                    }
                }
                continue;
            }

            // 命中了：合并 para 的 Run
            RunProperties? firstProps = null;
            foreach (var run in para.Elements<Run>())
            {
                if (firstProps == null && run.RunProperties != null)
                    firstProps = (RunProperties)run.RunProperties.CloneNode(true);
                para.RemoveChild(run);
            }
            var mergedRun = new Run();
            if (firstProps != null) mergedRun.RunProperties = firstProps;
            mergedRun.AppendChild(new Text(replaced) { Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(mergedRun);
        }
    }

    // ============================================================
    // 图片嵌入
    // ============================================================

    /// <summary>把含临时记号的段落替换为包含图片的 Run。</summary>
    private void InsertImageIntoParagraph(Paragraph para, string key, MainDocumentPart mainPart)
    {
        if (!_data.Images.TryGetValue(key, out var img)) return;

        // 创建新的 ImagePart 并写入字节
        var imagePart = mainPart.AddImagePart(GetImagePartType(img.ContentType));
        using (var stream = imagePart.GetStream(FileMode.Create))
        {
            stream.Write(img.Bytes, 0, img.Bytes.Length);
        }

        var relationshipId = mainPart.GetIdOfPart(imagePart);

        // 计算 EMU 尺寸（Word 使用 English Metric Unit，1 inch = 914400 EMU，1 cm = 360000 EMU）
        const int emuPerCm = 360000;
        var maxWidthCm = 15.5;  // A4 减去左右边距后的可用宽度
        var ratio = (double)img.PixelWidth / Math.Max(1, img.PixelHeight);
        var widthCm = Math.Min(maxWidthCm, img.PixelWidth / 96.0 * 2.54);
        var heightCm = widthCm / Math.Max(0.01, ratio);
        var widthEmu = (long)(widthCm * emuPerCm);
        var heightEmu = (long)(heightCm * emuPerCm);

        // 删除所有包含临时记号的 Run
        foreach (var run in para.Elements<Run>().ToList())
        {
            para.RemoveChild(run);
        }

        // 构造图片 Run 并追加到段落
        var drawing = BuildImageDrawing(relationshipId, img.SuggestedFileName, widthEmu, heightEmu, img.PixelWidth, img.PixelHeight);
        var imageRun = new Run(drawing);
        para.AppendChild(imageRun);
    }

    private static Drawing BuildImageDrawing(string relId, string name, long widthEmu, long heightEmu, int pixelW, int pixelH)
    {
        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = 1U, Name = name },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                EditId = "0",
            }
        );
    }

    private static PartTypeInfo GetImagePartType(string contentType)
    {
        return contentType switch
        {
            "image/png" => ImagePartType.Png,
            "image/jpeg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            "image/bmp" => ImagePartType.Bmp,
            _ => ImagePartType.Png,
        };
    }

    // ============================================================
    // 页面设置
    // ============================================================

    /// <summary>
    /// 应用页面设置：
    /// - UseTemplateSettings=true（默认）：跟随模板；
    ///   仅在用户显式指定了 PaperSize / Orientation / Margins 时，才覆盖模板里的对应项。
    /// - UseTemplateSettings=false：全部强制覆盖（缺项用默认值）。
    ///
    /// 规则：
    ///   写入 pgSz 时 **Width/Height 必须与 Orient 保持一致**，否则 Word/Spire.Doc
    ///   会读到"反方向"的页面尺寸。
    ///   - 纵向（Portrait）：w &lt; h（A4 → 11906 × 16838 twips）
    ///   - 横向（Landscape）：w &gt; h（A4 → 16838 × 11906 twips）
    /// </summary>
    private static void ApplyPageSettings(Body body, WordExportOptions options)
    {
        var sectionProps = body.Descendants<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.AppendChild(sectionProps);
        }

        var pgSz = sectionProps.Elements<PageSize>().FirstOrDefault();
        bool hadPgSz = pgSz != null;
        bool useTemplate = options.UseTemplateSettings;

        // ============================================================
        // 阶段 A：确定纸张基尺寸 (paperW, paperH)（纵向形式：w < h）
        // ============================================================
        uint paperW, paperH;

        if (options.PaperSize.HasValue)
        {
            var (w, h) = GetPaperDimensions(options.PaperSize.Value);
            paperW = w;
            paperH = h;
        }
        else if (hadPgSz && useTemplate)
        {
            // 从模板读取。如果模板方向是 landscape，交换回 (w<h) 形式
            uint tw, th;
            if (pgSz!.Width != null && pgSz.Height != null)
            {
                tw = pgSz.Width.Value;
                th = pgSz.Height.Value;
            }
            else
            {
                tw = 11906U; th = 16838U;
            }
            paperW = Math.Min(tw, th);
            paperH = Math.Max(tw, th);
        }
        else
        {
            paperW = 11906U; // 默认 A4 纵向
            paperH = 16838U;
        }

        // ============================================================
        // 阶段 B：确定最终方向 landscape
        // ============================================================
        bool landscape;
        if (options.Orientation.HasValue)
        {
            landscape = options.Orientation.Value == PageOrientation.Landscape;
        }
        else if (hadPgSz && useTemplate)
        {
            var orient = pgSz!.Orient?.Value;
            bool orientIsLandscape = orient == PageOrientationValues.Landscape;
            // 宽高比与 Orient 不一致时，以宽高比为准（有些模板不写 Orient 属性）
            bool ratioIsLandscape = (pgSz.Width?.Value ?? 0) > (pgSz.Height?.Value ?? 0);
            landscape = orientIsLandscape || ratioIsLandscape;
        }
        else
        {
            landscape = false; // 默认纵向
        }

        // ============================================================
        // 阶段 C：决定是否写入 PageSize.Width/Height/Orient
        // ============================================================
        bool mustWriteSize = !useTemplate
                             || options.PaperSize.HasValue
                             || options.Orientation.HasValue
                             || !hadPgSz;

        if (mustWriteSize)
        {
            uint finalW, finalH;
            if (landscape) { finalW = paperH; finalH = paperW; }
            else { finalW = paperW; finalH = paperH; }

            pgSz ??= new PageSize();
            if (!hadPgSz) sectionProps.PrependChild(pgSz);

            // Width/Height 与 Orient 必须一致写入（缺一不可）
            pgSz.Width = finalW;
            pgSz.Height = finalH;
            pgSz.Orient = landscape
                ? PageOrientationValues.Landscape
                : PageOrientationValues.Portrait;
        }

        // ============================================================
        // 阶段 D：页边距
        // ============================================================
        var pgMar = sectionProps.Elements<PageMargin>().FirstOrDefault();
        bool hadPgMar = pgMar != null;

        bool mustWriteMar = !useTemplate
                            || options.Margins != null
                            || !hadPgMar;

        if (!mustWriteMar) return;

        PageMargins margins = options.Margins ?? PageMargins.Default;
        var cmToTwip = (double cm) => (uint)(cm * 567);
        pgMar ??= new PageMargin();
        if (!hadPgMar) sectionProps.AppendChild(pgMar);
        pgMar.Top = (int)cmToTwip(margins.TopCm);
        pgMar.Bottom = (int)cmToTwip(margins.BottomCm);
        pgMar.Left = (uint)cmToTwip(margins.LeftCm);
        pgMar.Right = (uint)cmToTwip(margins.RightCm);
        pgMar.Header = (uint)cmToTwip(1.0);
        pgMar.Footer = (uint)cmToTwip(1.0);
        pgMar.Gutter = 0U;
    }

    private static (uint W, uint H) GetPaperDimensions(PaperSize paper)
    {
        return paper switch
        {
            PaperSize.A3 => (16838U, 23811U),         // 29.7 x 42 cm
            PaperSize.A4 => (11906U, 16838U),         // 21.0 x 29.7 cm
            PaperSize.A5 => (8391U, 11906U),          // 14.8 x 21 cm
            PaperSize.Letter => (12240U, 15840U),     // 8.5 x 11 in
            PaperSize.Legal => (12240U, 20160U),      // 8.5 x 14 in
            _ => (11906U, 16838U),
        };
    }
}

// DistinctBy Polyfill for older targets
internal static class LinqExt
{
    public static IEnumerable<T> DistinctBy<T, TKey>(this IEnumerable<T> src, System.Func<T, TKey> keySelector)
    {
        var seen = new HashSet<TKey>();
        foreach (var item in src)
        {
            if (seen.Add(keySelector(item))) yield return item;
        }
    }
}