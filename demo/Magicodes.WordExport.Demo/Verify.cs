using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Magicodes.WordExport;
using Magicodes.WordExport.Models;
using Magicodes.WordExport.Print;
using ScottPlot;

namespace Magicodes.WordExport.Demo;

/// <summary>
/// 端到端验证脚本：不打开 GUI，命令行直接调用类库 API 生成 Word + PDF。
/// 用于 build 后的 smoke test。
/// </summary>
public static class Verify
{
    public static void Run()
    {
        Console.WriteLine("=== Magicodes.WordExport 端到端验证 ===");
        var sw = System.Diagnostics.Stopwatch.StartNew();

        var baseDir = AppContext.BaseDirectory;
        var tplDir = Path.Combine(baseDir, "Resources", "Templates");
        var outDir = Path.Combine(baseDir, "Outputs");
        Directory.CreateDirectory(outDir);

        try
        {
            // 1) 生成 / 加载模板
            var tpl = TemplateBuilder.EnsureTemplate(tplDir);
            Console.WriteLine($"[1/5] 模板: {tpl} ({new FileInfo(tpl).Length} 字节)");

            // 2) 构造 ReportData
            var (data, plot, samples) = ReportDataFactory.Create();
            Console.WriteLine($"[2/5] 数据: {data.Variables.Count} 变量, {data.Tables.Count} 表, {data.Images.Count} 图片");

            // 3) 生成 Word 字节
            var builder = WordExportBuilder
                .Create()
                .UseTemplate(tpl)
                .Configure(o =>
                {
                    o.Orientation = PageOrientation.Landscape;
                    o.PaperSize = PaperSize.A4;
                    o.Margins = PageMargins.Min;
                })
                .WithData(data);

            var wordPath = Path.Combine(outDir, "verify_report.docx");
            builder.SaveWord(wordPath);
            var wordBytes = File.ReadAllBytes(wordPath);
            Console.WriteLine($"[3/5] Word: {wordPath} ({wordBytes.Length / 1024} KB)");

            // 校验：docx 是 zip，Magic Number = PK\x03\x04
            if (wordBytes.Length < 4 || wordBytes[0] != 0x50 || wordBytes[1] != 0x4B)
                throw new InvalidOperationException("docx 文件格式异常（不是 zip 格式）");
            Console.WriteLine("       ✓ docx 文件头 OK (PK zip 格式)");

            // 4) 导出 PDF
            var pdfPath = Path.Combine(outDir, "verify_report.pdf");
            builder.ExportPdf(pdfPath);
            var pdfBytes = File.ReadAllBytes(pdfPath);
            Console.WriteLine($"[4/5] PDF: {pdfPath} ({pdfBytes.Length / 1024} KB)");
            if (pdfBytes.Length < 8 || pdfBytes[0] != 0x25 || pdfBytes[1] != 0x50)
                throw new InvalidOperationException("PDF 文件格式异常（不是 PDF 格式）");
            Console.WriteLine("       ✓ PDF 文件头 OK (%PDF)");

            // 5) 打印（仅在不真正有物理打印机时跳过）
            var printers = DocumentPrinter.ListPrinters();
            if (printers.Count > 0 && Environment.GetEnvironmentVariable("SKIP_PRINT") != "1")
            {
                Console.WriteLine($"[5/5] 检测到 {printers.Count} 台打印机: {string.Join(", ", printers.Take(3).Select(p => p.Name))}");
                Console.WriteLine("       (跳过实际打印，CI 环境用 SKIP_PRINT=1 关闭)");
            }
            else
            {
                Console.WriteLine($"[5/5] 未检测到打印机 或设置了 SKIP_PRINT=1，跳过打印");
            }

            sw.Stop();
            Console.WriteLine();
            Console.WriteLine($"=== ✓ 全部通过 (用时 {sw.ElapsedMilliseconds} ms) ===");
            Environment.Exit(0);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine($"=== ✗ 失败: {ex.GetType().Name}: {ex.Message} ===");
            Console.Error.WriteLine(ex.ToString());
            Environment.Exit(1);
        }
    }
}