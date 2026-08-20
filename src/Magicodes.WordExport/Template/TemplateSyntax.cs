using System.Text.RegularExpressions;

namespace Magicodes.WordExport.Template;

/// <summary>
/// 类库支持的模板占位符语法常量。
///
/// 1) 变量：{{VarName}}    —— 普通变量替换
/// 2) 表格：{{Table>>TableName|RowNo}} ... {{FieldName|>>Table}} —— Magicodes.IE 风格表格
/// 3) 图片：@image#KEY:RELATIVE_PATH.png  —— 用户友好的内联图片占位符
///
/// 自定义图片语法对应当前用户截图里的 @image#1:Clipboard_Screenshot.png 写法。
/// </summary>
public static class TemplateSyntax
{
    /// <summary>@image#KEY:RELATIVE_PATH.EXT  —— KEY 是 ReportData.Images 中的键，PATH 是模板里的展示文本，会被替换成图片。</summary>
    public static readonly Regex ImagePlaceholderRegex =
        new(@"@image#(?<key>[A-Za-z0-9_\-]+)\s*:\s*(?<path>[^\s\r\n]+)", RegexOptions.Compiled);

    /// <summary>{{VarName}} 形式的变量占位符。</summary>
    public static readonly Regex VariableRegex =
        new(@"\{\{\s*(?<name>[A-Za-z0-9_\.]+)\s*\}\}", RegexOptions.Compiled);

    /// <summary>表格起始：{{Table>>TableName|RowNo}}。</summary>
    public static readonly Regex TableStartRegex =
        new(@"\{\{\s*Table\s*>>\s*(?<name>[A-Za-z0-9_\-]+)\s*\|?\s*(?<rowNo>[A-Za-z0-9_\-]*)\s*\}\}",
            RegexOptions.Compiled);

    /// <summary>表格字段：{{FieldName|>>Table}}。FieldName 对应行项的属性，|>>Table 是结束标记。</summary>
    public static readonly Regex TableFieldRegex =
        new(@"\{\{\s*(?<field>[A-Za-z0-9_\.]+)\s*\|\s*>>\s*Table\s*\}\}", RegexOptions.Compiled);

    /// <summary>内部临时记号：渲染过程中 @image 占位符会被替换为该记号，最后由 OpenXML 处理。</summary>
    public const string ImageSentinelPrefix = "<<<MAGICODES_WORD_EXPORT_IMAGE_";
    public const string ImageSentinelSuffix = ">>>";

    public static string WrapImageSentinel(string key) => $"{ImageSentinelPrefix}{key}{ImageSentinelSuffix}";

    public static bool TryUnwrapImageSentinel(string text, out string key)
    {
        if (text != null
            && text.StartsWith(ImageSentinelPrefix)
            && text.EndsWith(ImageSentinelSuffix))
        {
            key = text.Substring(ImageSentinelPrefix.Length,
                text.Length - ImageSentinelPrefix.Length - ImageSentinelSuffix.Length);
            return true;
        }
        key = string.Empty;
        return false;
    }
}