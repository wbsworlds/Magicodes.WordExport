using System;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Magicodes.WordExport.Pdf;

/// <summary>
/// 把 docx 字节导出为 PDF。底层用 FreeSpire.Doc（社区版免费，
/// 对单文档检测报告足够使用；超过 3 个表格 / 10 个段落会产生水印）。
/// </summary>
public sealed class PdfExporter
{
    /// <summary>
    /// 把 docx 字节转换为 PDF 字节。
    /// </summary>
    public byte[] ExportToPdfBytes(byte[] docxBytes, Models.WordExportOptions options)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("docx 字节为空", nameof(docxBytes));

        using var doc = new Document();
        using var inStream = new MemoryStream(docxBytes);
        doc.LoadFromStream(inStream, FileFormat.Docx2019);

        // 应用纸张大小与方向（再次确保打印侧与渲染侧一致）
        ApplyPaperSettings(doc, options);

        using var outStream = new MemoryStream();
        doc.SaveToStream(outStream, FileFormat.PDF);
        return outStream.ToArray();
    }

    /// <summary>
    /// 把 docx 字节直接保存为 PDF 文件。
    /// </summary>
    public string ExportToPdfFile(byte[] docxBytes, string targetPath, Models.WordExportOptions options)
    {
        var bytes = ExportToPdfBytes(docxBytes, options);
        File.WriteAllBytes(targetPath, bytes);
        return targetPath;
    }

    /// <summary>
    /// 应用纸张尺寸与方向。策略同 ApplyPageSettings：
    /// 默认 UseTemplateSettings=true 跟随模板，用户显式指定才覆盖。
    /// </summary>
    private static void ApplyPaperSettings(Document doc, Models.WordExportOptions options)
    {
        foreach (Spire.Doc.Section section in doc.Sections)
        {
            bool useTemplate = options.UseTemplateSettings;

            // 纸张
            if (options.PaperSize.HasValue || !useTemplate)
            {
                var size = options.PaperSize ?? Models.PaperSize.A4;
                section.PageSetup.PageSize = size switch
                {
                    Models.PaperSize.A3 => Spire.Doc.Documents.PageSize.A3,
                    Models.PaperSize.A4 => Spire.Doc.Documents.PageSize.A4,
                    Models.PaperSize.A5 => Spire.Doc.Documents.PageSize.A5,
                    Models.PaperSize.Letter => Spire.Doc.Documents.PageSize.Letter,
                    Models.PaperSize.Legal => Spire.Doc.Documents.PageSize.Legal,
                    _ => Spire.Doc.Documents.PageSize.A4,
                };
            }

            // 方向
            if (options.Orientation.HasValue || !useTemplate)
            {
                section.PageSetup.Orientation = (options.Orientation ?? Models.PageOrientation.Portrait) switch
                {
                    Models.PageOrientation.Landscape => Spire.Doc.Documents.PageOrientation.Landscape,
                    _ => Spire.Doc.Documents.PageOrientation.Portrait,
                };
            }

            // 页边距
            if (options.Margins != null || !useTemplate)
            {
                var m = options.Margins ?? Models.PageMargins.Default;
                section.PageSetup.Margins.Top = (float)m.TopCm * 28.35f;
                section.PageSetup.Margins.Bottom = (float)m.BottomCm * 28.35f;
                section.PageSetup.Margins.Left = (float)m.LeftCm * 28.35f;
                section.PageSetup.Margins.Right = (float)m.RightCm * 28.35f;
            }
        }
    }
}