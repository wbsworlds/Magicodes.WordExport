# Magicodes.WordExport

基于 **模板** 的 Word 报告生成库（.NET 8 / WPF）。用一套 docx 模板 + 占位符，配合流畅（Fluent）API，把变量、表格、图片、ScottPlot 图表渲染进 Word，并支持导出 PDF、打印预览与直接打印。

> 本项目是 Magicodes.IE.Word 的「模板渲染增强版」：不依赖 Magicodes.IE 的 DTO 标注模型，而是直接基于 `DocumentFormat.OpenXml` 操作 docx，支持更灵活的自由模板（变量 / 表格行复制 / 图片内联 / ScottPlot 图表）。

---

## 功能特性

| 类别 | 说明 |
| --- | --- |
| **模板渲染** | 基于 docx 模板渲染，保留原模板的样式、字体、表格边框 |
| **变量替换** | 段落 / 表格单元格中的 `{{VarName}}` 替换为变量值 |
| **动态表格** | 模板标记 `{{#TABLE_START:Name}}`，运行时按数据行数自动复制行并回填字段 |
| **图片内联** | `@image#KEY:filename.png` 占位符替换为真实图片（PNG / JPG / BMP / GIF） |
| **ScottPlot 图表** | 直接传入 `ScottPlot.Plot` 或配置 lambda，自动渲染为图片嵌入文档；内置中文字体支持，避免乱码 |
| **导出 PDF** | 通过 FreeSpire.Doc 将渲染结果导出为 PDF |
| **打印预览** | WPF 原生 `DocumentViewer` 分页预览（缩放 / 翻页 / 缩略图），无 Win32「不支持预览」提示 |
| **直接打印** | 自绘打印设置窗口（打印机 / 方向 / 份数 / 页码范围），调用系统打印管线 |
| **页面设置** | 纸张（A3/A4/A5/Letter/Legal）、方向（纵向/横向）、页边距；默认跟随模板，可逐项覆盖 |
| **流畅 API** | `WordExportBuilder.Create().UseTemplate()...Build()` 链式写法，支持同步 / 异步 |

---

## 环境要求

- **.NET 8 SDK**（Windows 平台，目标框架 `net8.0-windows`）
- Windows 桌面工作负载（WPF + Windows Forms）：`UseWPF=true`、`UseWindowsForms=true`
- 依赖库：`DocumentFormat.OpenXml` 3.0.2、`FreeSpire.Doc` 12.2.0、`Magicodes.IE.Word` 2.7.6、`ScottPlot` 5.0.55

> ⚠️ 本项目是 **Windows-only**：打印与图表预览依赖 WPF / GDI+，无法在 Linux / macOS 上运行 GUI 与打印相关功能。

---

## 目录结构

```
Magicodes.WordExport.sln
├── src/
│   └── Magicodes.WordExport/        # 核心类库
│       ├── WordExportBuilder.cs     # 流畅 API 主入口
│       ├── Models/                  # ReportData / WordExportOptions / RenderResult
│       ├── Render/TemplateRenderer.cs   # OpenXML 模板渲染核心
│       ├── Template/                # 占位符语法与预处理
│       ├── Pdf/PdfExporter.cs       # docx → PDF
│       ├── Print/                   # 打印 / 预览 / 设置窗口
│       └── ScottPlotFont.cs         # ScottPlot 中文字体扩展
└── demo/
    └── Magicodes.WordExport.Demo/   # 检测报告生成器（示例 WPF 应用）
        ├── MainWindow.cs            # DataGrid 编辑 + 图表预览 + 生成/导出/打印
        ├── ReportDataFactory.cs     # 示例数据（标准曲线）
        ├── TemplateBuilder.cs       # 运行时生成示例模板
        └── Verify.cs                # 命令行端到端 smoke test
```

---

## 模板语法

在 Word 里写好模板，用下列占位符标记动态内容：

### 1. 变量 `{{VarName}}`

```text
编号：{{ReportNo}}
检测员：{{Tester}}    复核员：{{Reviewer}}
```

### 2. 动态表格 `{{#TABLE_START:Name}}`

在表格的某个单元格内放起始标记，同表后续单元格放字段占位符 `{{FieldName}}`：

```text
{{#TABLE_START:Samples}}      ← 标记行（会被清掉，只用于定位）
{{Index}}                     ← 字段：按数据行自动复制
{{Name}}
{{Concentration}}
```

数据结构（每个元素是一个字段字典）：

```csharp
.AddTable("Samples", rows.Select(s => new Dictionary<string, object?>
{
    ["Index"] = s.Index,
    ["Name"]  = s.Name,
    ["Concentration"] = s.Concentration.ToString("F4"),
}))
```

### 3. 图片 `@image#KEY:filename.png`

```text
@image#1:curve.png           ← KEY=1，对应代码中的 AddImage/AddChart("1", ...)
```

---

## 快速开始

### 方式 A：引用源码（本仓库）

```bash
dotnet build Magicodes.WordExport.sln
```

在你的项目中引用 `src/Magicodes.WordExport/Magicodes.WordExport.csproj`。

### 方式 B：作为类库引用

把 `Magicodes.WordExport` 类库编译后的 dll 引用到你的 WPF 项目即可。

---

## API 用法

### 基础：渲染并保存 Word

```csharp
using Magicodes.WordExport;
using Magicodes.WordExport.Models;

var result = WordExportBuilder
    .Create()
    .UseTemplate("report_template.docx")          // 或 UseTemplate(byte[])
    .Configure(o =>
    {
        o.PaperSize = PaperSize.A4;
        o.Orientation = PageOrientation.Landscape; // 不设置则跟随模板
        o.Margins = PageMargins.Min;
    })
    .AddVariable("ReportNo", "RPT-001")
    .AddVariable("Tester", "张三")
    .AddTable("Samples", samples.Select(s => new Dictionary<string, object?>
    {
        ["Index"] = s.Index,
        ["Name"]  = s.Name,
        ["Abs"]   = s.Abs.ToString("F4"),
    }))
    .AddChart("1", p =>                          // 直接渲染 ScottPlot 图表
    {
        p.Add.Scatter(xs, ys);
        p.Title("标准曲线");
        p.XLabel("浓度 (mg/L)");
        p.YLabel("Abs");
    })
    .Build();

result.SaveWordAs("report.docx");
```

### 导出 PDF

```csharp
WordExportBuilder
    .Create()
    .UseTemplate("report_template.docx")
    .WithData(data)
    .ExportPdf("report.pdf");          // 同步
// .ExportPdfAsync("report.pdf")       // 异步
```

### 打印预览（WPF，需 STA 线程）

```csharp
var r = await WordExportBuilder
    .Create()
    .UseTemplate(tpl)
    .WithData(data)
    .PreviewAsync("检测报告");          // 打开自定义预览窗口
// r.IsSuccess / r.PaperCount
```

### 直接打印（自绘设置窗口）

```csharp
var r = await WordExportBuilder
    .Create()
    .UseTemplate(tpl)
    .WithData(data)
    .PrintDirectAsync("检测报告", ownerWindow);
```

### 渲染为字节后自行处理

```csharp
var result = WordExportBuilder.Create().UseTemplate(tpl).WithData(data).Build();
byte[] docx = result.WordBytes;          // 可直接写入 / 上传 / 流转
```

---

## 选项说明（`WordExportOptions`）

| 属性 | 默认 | 说明 |
| --- | --- | --- |
| `UseTemplateSettings` | `true` | `true`：仅当用户显式设置 PaperSize/Orientation/Margins 才覆盖模板；`false`：模板页面设置全部忽略 |
| `PaperSize` | `null` | `A3/A4/A5/Letter/Legal`，`null` 跟随模板 |
| `Orientation` | `null` | `Portrait/Landscape`，`null` 跟随模板 |
| `Margins` | `null` | 页边距（cm），`null` 跟随模板；`PageMargins.Default` / `PageMargins.Min` |
| `AddPageNumber` | `true` | 是否在页脚添加页码 |
| `PdfQuality` | `High` | PDF 导出质量（仅 FreeSpire.Doc 生效） |
| `Warnings` | — | 渲染过程收集的非致命警告 |

---

## ScottPlot 图表与中文字体

图表统一通过 `ScottPlotFont.ApplyCjkFont` 应用 `Microsoft YaHei UI` 字体，解决 headless 渲染（`Plot.GetImage`）时中文标签乱码问题：

```csharp
plot.ApplyCjkFont(bold: true);   // 标题 / 轴标签 / 图例全局加粗 + 中文字体
```

- 传入 `Plot` 对象：`AddChart("1", plot)`
- 传入配置 lambda（更简洁）：`AddChart("1", p => { p.Add.Scatter(xs, ys); })`

---

## 打印与预览实现要点

- **预览**：使用 WPF 原生 `DocumentViewer` 承载 GDI 分页图像，外观专业、无第三方依赖；打印按钮被接管为自绘设置窗口，规避 Win32「此应用不支持打印预览」提示。
- **分页渲染**：`FreeSpire.Doc` 加载 docx → `PreviewPrintController` 输出每页 `Image` → 转为 `FixedDocument`。
- **设置窗口**：自绘的打印机 / 方向 / 份数 / 页码范围选择，借用 WPF `PrintDialog.PrintDocument` API 投递作业（不弹系统对话框）。
- **旧版 WinForms 预览**仍可用：`DocumentPrinter.ShowPreview(...)`（`PrintPreviewDialog`）。

---

## 运行 Demo（检测报告生成器）

```bash
dotnet run --project demo/Magicodes.WordExport.Demo
```

窗体功能：编辑样本 DataGrid → 实时刷新标准曲线图 → 一键「生成 Word / 导出 PDF / 打印预览 / 打印」。模板在首次运行时由 `TemplateBuilder.EnsureTemplate` 自动生成到 `Resources/Templates/`。

### 命令行端到端验证（CI / smoke test）

```bash
dotnet run --project demo/Magicodes.WordExport.Demo -- --verify
# 或跳过实际打印（无打印机 / CI 环境）：
SKIP_PRINT=1 dotnet run --project demo/Magicodes.WordExport.Demo -- --verify
```

验证输出 Word + PDF 文件头是否正确，并列出本机打印机。

---

## 构建与测试

```bash
dotnet build Magicodes.WordExport.sln      # 类库 + Demo，要求 0 error
```

构建产物与运行期生成的模板 / 输出目录已在 `.gitignore` 中忽略，仓库只保留源码与文档。

---

## 已知限制

1. **Windows-only**：打印 / 预览依赖 WPF 与 GDI+，非 Windows 平台仅可用纯渲染 / 导出字节逻辑（需自行剥离打印相关调用）。
2. **FreeSpire.Doc 水印**：社区版对超过 3 个表格 / 10 个段落的文档导出 PDF 会带水印；生产环境建议替换为 Spire.Doc 商业版或 LibreOffice 命令行转换。
3. **图片尺寸**：内联图片最大宽度按 A4 可用宽度（约 15.5cm）等比缩放。
4. **变量命名**：占位符仅支持 `[A-Za-z0-9_.\-]`，且区分大小写不敏感（字典以 `OrdinalIgnoreCase` 比较）。

---

## 技术栈

- 语言：C# 12 / .NET 8（WPF + Windows Forms）
- 模板：DocumentFormat.OpenXml 3.0.2
- 图表：ScottPlot 5.0.55
- PDF：FreeSpire.Doc 12.2.0
- 报表框架参考：Magicodes.IE.Word 2.7.6

---

## 许可证

[MIT](LICENSE) © 2026 Wang Bo
