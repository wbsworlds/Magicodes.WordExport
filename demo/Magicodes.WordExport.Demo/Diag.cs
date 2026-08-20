using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Magicodes.WordExport.Demo;

/// <summary>诊断工具：读取生成的 docx，打印 SectionProperties 实际值（用于调试页面方向）。</summary>
public static class Diag
{
    public static string DumpDocxPageInfo(string docxPath)
    {
        var sw = new System.Text.StringBuilder();
        sw.AppendLine("========== docx SectionProperties dump ==========");
        try
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            using var pkg = Package.Open(docxPath, FileMode.Open, FileAccess.Read);
            var docPart = pkg.GetParts()
                .FirstOrDefault(p => p.Uri.OriginalString.EndsWith("/document.xml", StringComparison.OrdinalIgnoreCase));
            if (docPart == null) { sw.AppendLine("  (document.xml 未找到)"); return sw.ToString(); }

            using var stream = docPart.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = XmlReader.Create(stream);
            var xdoc = XDocument.Load(reader);
            var sectPrs = xdoc.Descendants(w + "sectPr").ToList();

            for (int i = 0; i < sectPrs.Count; i++)
            {
                var s = sectPrs[i];
                var pgSz = s.Element(w + "pgSz");
                var pgMar = s.Element(w + "pgMar");
                if (pgSz != null)
                {
                    var wA = pgSz.Attribute(w + "w");
                    var hA = pgSz.Attribute(w + "h");
                    var oA = pgSz.Attribute(w + "orient");
                    bool okW = double.TryParse(wA?.Value, out var wv);
                    bool okH = double.TryParse(hA?.Value, out var hv);
                    string? info = null;
                    if (okW && okH)
                    {
                        info = $"  [{i}] pgSz: w={wA.Value} h={hA.Value} orient={oA?.Value ?? "(portrait)"}"
                               + $"  ratio={(wv / hv):F3}"
                               + $"  => {(wv > hv ? "LANDSCAPE" : "PORTRAIT")}";
                    }
                    else { info = $"  [{i}] pgSz: w/h 缺失"; }
                    sw.AppendLine(info);
                }
                if (pgMar != null)
                {
                    sw.AppendLine($"  [{i}] pgMar: top={pgMar.Attribute(w + "top")?.Value}"
                                  + $" bottom={pgMar.Attribute(w + "bottom")?.Value}"
                                  + $" left={pgMar.Attribute(w + "left")?.Value}"
                                  + $" right={pgMar.Attribute(w + "right")?.Value}");
                }
            }
        }
        catch (Exception ex)
        {
            sw.AppendLine("  读取失败: " + ex.Message);
        }
        sw.AppendLine("================================================");
        return sw.ToString();
    }
}