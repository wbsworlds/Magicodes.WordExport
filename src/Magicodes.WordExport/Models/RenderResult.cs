using System;
using System.IO;

namespace Magicodes.WordExport.Models;

/// <summary>
/// 渲染结果。
/// </summary>
public sealed class RenderResult
{
    /// <summary>最终 Word 文档字节（.docx）。</summary>
    public byte[] WordBytes { get; }

    /// <summary>PDF 字节（如果调用过 ExportPdf/Print 时生成）。</summary>
    public byte[]? PdfBytes { get; set; }

    /// <summary>渲染过程中产生的非致命警告。</summary>
    public System.Collections.Generic.List<string> Warnings { get; } = new();

    /// <summary>本次渲染输出目录；其中的 .docx 是临时文件，可用于打印或预览。</summary>
    public string OutputDirectory { get; }

    public RenderResult(byte[] wordBytes, string outputDirectory)
    {
        WordBytes = wordBytes ?? throw new ArgumentNullException(nameof(wordBytes));
        OutputDirectory = outputDirectory ?? throw new ArgumentNullException(nameof(outputDirectory));
    }

    /// <summary>把 Word 字节写到磁盘。</summary>
    public string SaveWordAs(string path)
    {
        File.WriteAllBytes(path, WordBytes);
        return path;
    }

    /// <summary>把 PDF 字节写到磁盘。</summary>
    public string SavePdfAs(string path)
    {
        if (PdfBytes == null)
            throw new InvalidOperationException("PdfBytes 为空。请先调用 ExportPdf 或 Print。");
        File.WriteAllBytes(path, PdfBytes);
        return path;
    }
}