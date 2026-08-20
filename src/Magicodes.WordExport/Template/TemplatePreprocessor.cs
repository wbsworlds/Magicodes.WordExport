using System.Collections.Generic;
using System.Text.RegularExpressions;
using Magicodes.WordExport.Models;
using Magicodes.WordExport.Template;

namespace Magicodes.WordExport.Template;

/// <summary>
/// 把 ReportData 中的占位符展开成可识别的记号 / 字面量，
/// 供 <see cref="Render.TemplateRenderer"/> 在 OpenXML 层处理。
///
/// 这一步只做字符串处理（不动 docx 结构），保证后续 docx 操作是确定性的。
/// </summary>
public static class TemplatePreprocessor
{
    /// <summary>
    /// 把单个 Run 内的纯文本进行替换：
    ///   @image#KEY:filename.png -> &lt;&lt;&lt;MAGICODES_WORD_EXPORT_IMAGE_KEY&gt;&gt;&gt;
    ///   {{VarName}} -> 值（占位符，渲染阶段再回填；这里仅做语法识别，保留为变量名）
    /// </summary>
    /// <returns>替换后的文本和是否发现图片占位符。</returns>
    public static PreprocessResult PreprocessRunText(string raw, ReportData data)
    {
        var foundImage = false;
        var keysFound = new List<string>();

        // 1) 图片占位符
        var imgReplaced = TemplateSyntax.ImagePlaceholderRegex.Replace(raw, m =>
        {
            var key = m.Groups["key"].Value;
            if (!data.Images.ContainsKey(key))
            {
                // 未注册的图片占位符：保留原样以便用户排查
                return m.Value;
            }
            foundImage = true;
            keysFound.Add(key);
            return TemplateSyntax.WrapImageSentinel(key);
        });

        // 2) 变量占位符：仅当变量存在时才替换；否则保留原样
        var varReplaced = TemplateSyntax.VariableRegex.Replace(imgReplaced, m =>
        {
            var name = m.Groups["name"].Value;
            // 排除表格开始标记 {{Table>>Name|RowNo}} —— 它们是结构标记，由渲染器处理
            if (name.StartsWith("Table", System.StringComparison.OrdinalIgnoreCase)
                && m.Value.Contains(">>"))
            {
                return m.Value;
            }
            if (data.Variables.TryGetValue(name, out var val))
            {
                return val?.ToString() ?? string.Empty;
            }
            return m.Value;
        });

        return new PreprocessResult(varReplaced, foundImage, keysFound);
    }
}

public readonly record struct PreprocessResult(string Text, bool FoundImage, IReadOnlyList<string> ImageKeys);