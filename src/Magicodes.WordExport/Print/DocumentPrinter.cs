using System;
using System.IO;
using Spire.Doc;
using Spire.Doc.Printing;

namespace Magicodes.WordExport.Print;

/// <summary>
/// 文档打印。基于 FreeSpire.Doc 把 docx 加载到内存并交给 System.Drawing.Printing.PrintDocument 打印。
/// </summary>
public sealed class DocumentPrinter
{
    /// <summary>
    /// 调用默认打印机直接打印。
    /// </summary>
    public void Print(byte[] docxBytes, Models.WordExportOptions options)
    {
        using var doc = new Document();
        using var ms = new MemoryStream(docxBytes);
        doc.LoadFromStream(ms, Spire.Doc.FileFormat.Docx2019);

        using var printDoc = doc.PrintDocument;
        printDoc.Print();
    }

    /// <summary>
    /// 打印到指定打印机。
    /// </summary>
    public void PrintOn(byte[] docxBytes, Models.WordExportOptions options, string printerName, int copies = 1)
    {
        if (string.IsNullOrWhiteSpace(printerName))
            throw new ArgumentException("打印机名不能为空", nameof(printerName));

        using var doc = new Document();
        using var ms = new MemoryStream(docxBytes);
        doc.LoadFromStream(ms, Spire.Doc.FileFormat.Docx2019);

        using var printDoc = doc.PrintDocument;
        printDoc.PrinterSettings.PrinterName = printerName;
        printDoc.PrinterSettings.Copies = (short)Math.Max(1, copies);
        printDoc.Print();
    }

    /// <summary>
    /// 显示打印预览窗口（WinForms PrintPreviewDialog）。
    /// 预览窗口工具栏自带"打印"按钮，所见即所得。
    /// </summary>
    public void ShowPreview(byte[] docxBytes, Models.WordExportOptions options)
    {
        using var printDoc = CreatePrintDocument(docxBytes);
        using var dialog = new System.Windows.Forms.PrintPreviewDialog
        {
            Document = printDoc,
            Width = 1000,
            Height = 760,
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
            UseAntiAlias = true,
        };
        dialog.ShowDialog();
    }

    /// <summary>
    /// 创建 PrintDocument（预览/打印共用）。
    /// WPF 项目可以拿它自己接 DocumentViewer 或 XPS 流水线，不必依赖 WinForms 对话框。
    /// 调用方负责 Dispose。
    /// </summary>
    public System.Drawing.Printing.PrintDocument CreatePrintDocument(byte[] docxBytes)
    {
        var doc = new Document();
        var ms = new MemoryStream(docxBytes);
        doc.LoadFromStream(ms, Spire.Doc.FileFormat.Docx2019);
        return doc.PrintDocument;
    }

    /// <summary>
    /// 枚举本机可用的打印机列表（用于 demo UI 选择）。
    /// </summary>
    public static System.Collections.Generic.List<PrinterInfo> ListPrinters()
    {
        var list = new System.Collections.Generic.List<PrinterInfo>();
        foreach (string name in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
        {
            list.Add(new PrinterInfo(name));
        }
        return list;
    }
}

public sealed record PrinterInfo(string Name);